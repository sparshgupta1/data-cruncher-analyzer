using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Collections;

/// <summary>
/// Pivots the data
/// </summary>
public class Pivot
{
    public DataTable _SourceTable = new DataTable();
    private IEnumerable<DataRow> _Source = new List<DataRow>();

    public DataTable _SourceTableBkup = new DataTable(); //Used as a backup for _SourceTable when it is modified

    public DataTable _SpecTable = new DataTable();
    private IEnumerable<DataRow> _Spec = new List<DataRow>();

    private string _errormsg;
    public string errormsg { get { return _errormsg; } set { _errormsg = value; } }

    public Pivot(DataTable SourceTable, DataTable SpecTable = null)
    {
        _SourceTable = SourceTable;
        _Source = SourceTable.Rows.Cast<DataRow>();

        _SourceTableBkup = _SourceTable;

        if (_SpecTable != null) _SpecTable = SpecTable;
        if (_SpecTable != null) _Spec = SpecTable.Rows.Cast<DataRow>();
    }
    /// <summary>
    /// Prints the DataTable
    /// </summary>
    /// <param name="dt"></param>
    public void printDataTable(DataTable dt)
    {
        for (int i = 0; i < dt.Columns.Count; i++)
        {
            System.Diagnostics.Debug.Write(dt.Columns[i].ToString() + " | ");
        }
        foreach (DataRow dataRow in dt.Rows)
        {
            string LineStr = "";
            foreach (var item in dataRow.ItemArray)
            {
                LineStr = LineStr + " | " + item.ToString();
            }
            System.Diagnostics.Debug.WriteLine(LineStr);
        }
    }
    public DataTable CreateSpec(string[] DataFields, AggregateFunction[] Aggregate, string[] RowFields, string[] ColumnFields, double NSigma = 3)
    {
        DataTable dt = new DataTable();
        try
        {
            var RowList = _SourceTable.DefaultView.ToTable(true, RowFields.Concat(ColumnFields).ToArray()).AsEnumerable().ToList();
            for (int index = RowFields.Count() - 1; index >= 0; index--)
                RowList = RowList.OrderBy(x => x.Field<object>(RowFields[index])).ToList();

            //dt.Columns.Add(RowFields);
            foreach (string s in RowFields)
                dt.Columns.Add(s);

            dt.Columns.Add("null"); //Extra column to seperate the input and output

            Hashtable FunHash = new Hashtable();
            foreach (var fun in Aggregate)
                FunHash[fun] = "_" + fun.ToString();
            FunHash[AggregateFunction.Average] = "_Typical";

            foreach (string Data in DataFields)
            {
                foreach (var fun in Aggregate)
                    if (!fun.ToString().Contains("Spec")) dt.Columns.Add(Data + FunHash[fun].ToString());  // Cretes the result columns.//
                dt.Columns.Add(Data + "_Delta"); //Correction factor
            }

            foreach (var RowName in RowList)
            {
                DataRow row = dt.NewRow();
                string strFilter = string.Empty;
                string strSpecFilter = string.Empty;

                foreach (string Field in RowFields)
                {
                    row[Field] = RowName[Field];
                    strFilter += " and " + Field + " = '" + RowName[Field].ToString() + "'";
                }
                strFilter = strFilter.Substring(5);

                foreach (var fun in Aggregate)
                {
                    string filter = strFilter;
                    foreach (string Data in DataFields)
                        if (!fun.ToString().Contains("Spec")) row[Data + FunHash[fun].ToString()] = GetData(new string[] { filter }, Data, fun);
                }
                dt.Rows.Add(row);
            }
        }
        catch (Exception ex) { errormsg = errormsg + "," + ex.Message; }
        return dt;
    }

    public DataTable ManagedPivotData(string DataField, AggregateFunction[] Aggregate, string[] Iterator, string[] RowFields, string[] ColumnFields, double NSigma = 3)
    {
        DataTable dt = new DataTable();
        try
        {
            string Separator = "~";
            var RowList = _SourceTable.DefaultView.ToTable(true, Iterator).AsEnumerable().ToList();
            for (int index = Iterator.Count() - 1; index >= 0; index--)
                RowList = RowList.OrderBy(x => x.Field<object>(Iterator[index])).ToList();

            var RowList_RowFields = _SourceTable.DefaultView.ToTable(true, RowFields).AsEnumerable().ToList();
            for (int index = RowFields.Count() - 1; index >= 0; index--)
                RowList_RowFields = RowList_RowFields.OrderBy(x => x.Field<object>(RowFields[index])).ToList();

            // Gets the list of columns .(dot) separated.
            var ColListFormula = (from x in _SourceTable.AsEnumerable()
                                  select new
                                  {
                                      Name = ColumnFields.Select(n => x.Field<object>(n))
                                          .Aggregate((a, b) => a += Separator + b.ToString())
                                  })
                               .Distinct()
                               .OrderBy(m => m.Name);

            List<string> ColList = new List<string>();
            if (!(ColumnFields == null)) if (ColumnFields.Length > 0) foreach (var col in ColListFormula) ColList.Add(col.Name.ToString());

            //Include Spec filter
            DataColumnCollection Speccolumns = _SpecTable.Columns;

            bool ColumnFiledHasSpec = false;
            foreach (string col in ColumnFields) if (Speccolumns.Contains(col)) ColumnFiledHasSpec = true;

            //Add a column for Parameter and Unit
            dt.Columns.Add("Parameter");

            //dt.Columns.Add(RowFields);
            foreach (string s in Iterator)
                dt.Columns.Add(s);

            bool SpecNeeded = false;
            foreach (var fun in Aggregate) if (fun.ToString().Contains("Spec")) SpecNeeded = true;
            if ((!ColumnFiledHasSpec) & SpecNeeded)
            {
                if (Aggregate.Contains(AggregateFunction.SpecMin)) dt.Columns.Add("SpecMin");
                if (Aggregate.Contains(AggregateFunction.SpecMax)) dt.Columns.Add("SpecMax");
                if (Aggregate.Contains(AggregateFunction.SpecTypical)) dt.Columns.Add("SpecTypical");
            }

            if (ColumnFields.Length > 0)
                foreach (string col in ColList)
                {
                    foreach (var fun in Aggregate)
                        if (!fun.ToString().Contains("Spec")) dt.Columns.Add(col + '#' + fun.ToString());  // Cretes the result columns.//
                }
            else
            {
                foreach (var fun in Aggregate)
                    if (!fun.ToString().Contains("Spec")) dt.Columns.Add(fun.ToString());  // Cretes the result columns.//
            }
            if (Iterator.Length > 0)
            {
                foreach (var RowName in RowList)
                {
                    DataRow row = dt.NewRow();
                    row["Parameter"] = DataField;
                    string strPrimaryFilter = string.Empty;
                    string strSpecFilter = string.Empty;
                    List<string> strSecondaryFilters = new List<string>();

                    foreach (string Field in Iterator)
                    {
                        row[Field] = RowName[Field];
                        strPrimaryFilter += " and " + Field + " = '" + RowName[Field].ToString() + "'";

                        if (Speccolumns.Contains(Field)) strSpecFilter += " and " + Field + " = '" + RowName[Field].ToString() + "'";
                    }
                    strPrimaryFilter = strPrimaryFilter.Substring(5);

                    foreach (var r in RowList_RowFields)
                    {
                        string tempstrSecondaryFilter = string.Empty;
                        foreach (string Field
                            in RowFields.Except(Iterator).ToArray())
                        {
                            tempstrSecondaryFilter += " and " + Field + " = '" + r[Field].ToString() + "'";
                            if (Speccolumns.Contains(Field)) strSpecFilter += " and " + Field + " = '" + r[Field].ToString() + "'";
                        }
                        strSecondaryFilters.Add(strPrimaryFilter + " and " + tempstrSecondaryFilter.Substring(5));
                    }
                    strSecondaryFilters = strSecondaryFilters.Distinct().ToList<string>();
                    if (strSpecFilter.Length > 5) strSpecFilter = strSpecFilter.Substring(5);

                    if (ColumnFields.Length > 0)
                    {
                        if ((!ColumnFiledHasSpec) & SpecNeeded)
                        {
                            string filter = strSpecFilter;
                            if (Aggregate.Contains(AggregateFunction.SpecMin)) row["SpecMin"] = IncludeSpec(filter, DataField, "SpecMin");
                            if (Aggregate.Contains(AggregateFunction.SpecMax)) row["SpecMax"] = IncludeSpec(filter, DataField, "SpecMax");
                            if (Aggregate.Contains(AggregateFunction.SpecTypical)) row["SpecTypical"] = IncludeSpec(filter, DataField, "SpecTypical");
                        }

                        foreach (string col in ColList)
                        {
                            string filter = strPrimaryFilter;
                            string[] strColValues = col.Split(Separator.ToCharArray(), StringSplitOptions.None);
                            for (int i = 0; i < ColumnFields.Length; i++)
                                filter += " and " + ColumnFields[i] + " = '" + strColValues[i] + "'";
                            foreach (var fun in Aggregate)
                            {
                                if (!fun.ToString().Contains("Spec")) row[col + '#' + fun.ToString()] = GetData(new string[] { filter }, DataField, fun, NSigma);
                                else
                                {
                                    //filter = strSpecFilter;
                                    //row[col + '#' + fun.ToString()] = IncludeSpec(filter, DataField, fun.ToString());
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (var fun in Aggregate)
                        {
                            string filter = string.Empty;
                            if (fun != AggregateFunction.NSigmaNegMin && fun != AggregateFunction.NSigmaPosMax && fun != AggregateFunction.NSigmaNegCorrespondingStdDev && fun != AggregateFunction.NSigmaPosCorrespondingStdDev)
                            {
                                filter = strPrimaryFilter;
                                if (!fun.ToString().Contains("Spec")) row[fun.ToString()] = GetData(new string[] { filter }, DataField, fun, NSigma);
                                else
                                {
                                    filter = strSpecFilter;
                                    row[fun.ToString()] = IncludeSpec(filter, DataField, fun.ToString());
                                }
                            }
                            else
                            {
                                if (!fun.ToString().Contains("Spec")) row[fun.ToString()] = GetData(strSecondaryFilters.Distinct().ToArray(), DataField, fun, NSigma);
                                else
                                {
                                    filter = strSpecFilter;
                                    row[fun.ToString()] = IncludeSpec(filter, DataField, fun.ToString());
                                }

                            }
                        }
                    }
                    dt.Rows.Add(row);
                } // foreach (var RowName in RowList)               
            } //if (RowFields.Length > 0)
            else
            {
                DataRow row = dt.NewRow();
                string strFilter = string.Empty;
                string strSpecFilter = string.Empty;

                if ((!ColumnFiledHasSpec) & SpecNeeded)
                {
                    string filter = strSpecFilter;
                    if (Aggregate.Contains(AggregateFunction.SpecMin)) row["SpecMin"] = IncludeSpec(filter, DataField, "SpecMin");
                    if (Aggregate.Contains(AggregateFunction.SpecMax)) row["SpecMax"] = IncludeSpec(filter, DataField, "SpecMax");
                    if (Aggregate.Contains(AggregateFunction.SpecTypical)) row["SpecTypical"] = IncludeSpec(filter, DataField, "SpecTypical");
                }

                foreach (string col in ColList)
                {
                    string filter = strFilter;
                    string[] strColValues = col.Split(Separator.ToCharArray(), StringSplitOptions.None);
                    for (int i = 0; i < ColumnFields.Length; i++)
                        filter += " and " + ColumnFields[i] + " = '" + strColValues[i] + "'";

                    filter = filter.Substring(5);
                    if (strSpecFilter.Length > 5) strSpecFilter = strSpecFilter.Substring(5);

                    foreach (var fun in Aggregate)
                    {
                        if (!fun.ToString().Contains("Spec")) row[col + '#' + fun.ToString()] = GetData(new string[] { filter }, DataField, fun, NSigma);
                        else
                        {
                            //filter = strSpecFilter;
                            //row[col + '#' + fun.ToString()] = IncludeSpec(filter, DataField, fun.ToString());
                        }
                    }
                }
                dt.Rows.Add(row);
            }
        }
        catch (Exception ex) { errormsg = errormsg + "," + ex.Message; }
        return dt;
    }

    public DataTable PivotData(string DataField, AggregateFunction[] Aggregate, string[] RowFields, string[] ColumnFields, double NSigma = 3)
    {
        DataTable dt = new DataTable();
        try
        {
            string Separator = "~";
            var RowList = _SourceTable.DefaultView.ToTable(true, RowFields).AsEnumerable().ToList();
            for (int index = RowFields.Count() - 1; index >= 0; index--)
                RowList = RowList.OrderBy(x => x.Field<object>(RowFields[index])).ToList();
            // Gets the list of columns .(dot) separated.
            var ColListFormula = (from x in _SourceTable.AsEnumerable()
                                  select new
                                  {
                                      Name = ColumnFields.Select(n => x.Field<object>(n))
                                          .Aggregate((a, b) => a += Separator + b.ToString())
                                  })
                               .Distinct()
                               .OrderBy(m => m.Name);

            List<string> ColList = new List<string>();
            if (!(ColumnFields == null)) if (ColumnFields.Length > 0) foreach (var col in ColListFormula) ColList.Add(col.Name.ToString());

            //Include Spec filter
            DataColumnCollection Speccolumns = _SpecTable.Columns;

            bool ColumnFiledHasSpec = false;
            foreach (string col in ColumnFields) if (Speccolumns.Contains(col)) ColumnFiledHasSpec = true;

            //Add a column for Parameter and Unit
            dt.Columns.Add("Parameter");

            //dt.Columns.Add(RowFields);
            foreach (string s in RowFields)
                dt.Columns.Add(s);

            bool SpecNeeded = false;
            foreach (var fun in Aggregate) if (fun.ToString().Contains("Spec")) SpecNeeded = true;
            if ((!ColumnFiledHasSpec) & SpecNeeded)
            {
                if (Aggregate.Contains(AggregateFunction.SpecMin)) dt.Columns.Add("SpecMin");
                if (Aggregate.Contains(AggregateFunction.SpecMax)) dt.Columns.Add("SpecMax");
                if (Aggregate.Contains(AggregateFunction.SpecTypical)) dt.Columns.Add("SpecTypical");
            }

            if (ColumnFields.Length > 0)
                foreach (string col in ColList)
                {
                    foreach (var fun in Aggregate)
                        if (!fun.ToString().Contains("Spec")) dt.Columns.Add(col + '#' + fun.ToString());  // Cretes the result columns.//
                }
            else
            {
                foreach (var fun in Aggregate)
                    if (!fun.ToString().Contains("Spec")) dt.Columns.Add(fun.ToString());  // Cretes the result columns.//
            }
            if (RowFields.Length > 0)
            {
                foreach (var RowName in RowList)
                {
                    DataRow row = dt.NewRow();
                    row["Parameter"] = DataField;
                    string strFilter = string.Empty;
                    string strSpecFilter = string.Empty;

                    foreach (string Field in RowFields)
                    {
                        row[Field] = RowName[Field];
                        strFilter += " and " + Field + " = '" + RowName[Field].ToString() + "'";

                        if (Speccolumns.Contains(Field)) strSpecFilter += " and " + Field + " = '" + RowName[Field].ToString() + "'";
                    }
                    strFilter = strFilter.Substring(5);
                    if (strSpecFilter.Length > 5) strSpecFilter = strSpecFilter.Substring(5);

                    if (ColumnFields.Length > 0)
                    {
                        if ((!ColumnFiledHasSpec) & SpecNeeded)
                        {
                            string filter = strSpecFilter;
                            if (Aggregate.Contains(AggregateFunction.SpecMin)) row["SpecMin"] = IncludeSpec(filter, DataField, "SpecMin");
                            if (Aggregate.Contains(AggregateFunction.SpecMax)) row["SpecMax"] = IncludeSpec(filter, DataField, "SpecMax");
                            if (Aggregate.Contains(AggregateFunction.SpecTypical)) row["SpecTypical"] = IncludeSpec(filter, DataField, "SpecTypical");
                        }

                        foreach (string col in ColList)
                        {
                            string filter = strFilter;
                            string[] strColValues = col.Split(Separator.ToCharArray(), StringSplitOptions.None);
                            for (int i = 0; i < ColumnFields.Length; i++)
                                filter += " and " + ColumnFields[i] + " = '" + strColValues[i] + "'";
                            foreach (var fun in Aggregate)
                            {
                                if (!fun.ToString().Contains("Spec")) row[col + '#' + fun.ToString()] = GetData(new string[] { filter }, DataField, fun, NSigma);
                                else
                                {
                                    //filter = strSpecFilter;
                                    //row[col + '#' + fun.ToString()] = IncludeSpec(filter, DataField, fun.ToString());
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (var fun in Aggregate)
                        {
                            string filter = strFilter;
                            if (!fun.ToString().Contains("Spec")) row[fun.ToString()] = GetData(new string[] { filter }, DataField, fun, NSigma);
                            else
                            {
                                filter = strSpecFilter;
                                row[fun.ToString()] = IncludeSpec(filter, DataField, fun.ToString());
                            }
                        }
                    }
                    dt.Rows.Add(row);
                } // foreach (var RowName in RowList)
            } //if (RowFields.Length > 0)
            else
            {
                DataRow row = dt.NewRow();
                string strFilter = string.Empty;
                string strSpecFilter = string.Empty;

                if ((!ColumnFiledHasSpec) & SpecNeeded)
                {
                    string filter = strSpecFilter;
                    if (Aggregate.Contains(AggregateFunction.SpecMin)) row["SpecMin"] = IncludeSpec(filter, DataField, "SpecMin");
                    if (Aggregate.Contains(AggregateFunction.SpecMax)) row["SpecMax"] = IncludeSpec(filter, DataField, "SpecMax");
                    if (Aggregate.Contains(AggregateFunction.SpecTypical)) row["SpecTypical"] = IncludeSpec(filter, DataField, "SpecTypical");
                }

                foreach (string col in ColList)
                {
                    string filter = strFilter;
                    string[] strColValues = col.Split(Separator.ToCharArray(), StringSplitOptions.None);
                    for (int i = 0; i < ColumnFields.Length; i++)
                        filter += " and " + ColumnFields[i] + " = '" + strColValues[i] + "'";

                    filter = filter.Substring(5);
                    if (strSpecFilter.Length > 5) strSpecFilter = strSpecFilter.Substring(5);

                    foreach (var fun in Aggregate)
                    {
                        if (!fun.ToString().Contains("Spec")) row[col + '#' + fun.ToString()] = GetData(new string[] { filter }, DataField, fun, NSigma);
                        else
                        {
                            //filter = strSpecFilter;
                            //row[col + '#' + fun.ToString()] = IncludeSpec(filter, DataField, fun.ToString());
                        }
                    }
                }
                dt.Rows.Add(row);
            }
        }
        catch (Exception ex) { errormsg = errormsg + "," + ex.Message; }
        return dt;
    }
    public DataTable FilteredPivotData(string DataField, AggregateFunction[] Aggregate, string[] RowFields, string[] ColumnFields, double NSigma = 3, string AddFilter = null)
    {
        DataTable dt = new DataTable();
        try
        {
            string Separator = "$#";
            var RowList = _SourceTable.DefaultView.ToTable(true, RowFields).AsEnumerable().ToList();
            for (int index = RowFields.Count() - 1; index >= 0; index--)
                RowList = RowList.OrderBy(x => x.Field<object>(RowFields[index])).ToList();
            // Gets the list of columns .(dot) separated.
            var ColListFormula = (from x in _SourceTable.AsEnumerable()
                                  select new
                                  {
                                      Name = ColumnFields.Select(n => x.Field<object>(n))
                                          .Aggregate((a, b) => a += Separator + b.ToString())
                                  })
                               .Distinct()
                               .OrderBy(m => m.Name);

            List<string> ColList = new List<string>();
            if (!(ColumnFields == null)) if (ColumnFields.Length > 0) foreach (var col in ColListFormula) ColList.Add(col.Name.ToString());

            //Include Spec filter
            DataColumnCollection Speccolumns = _SpecTable.Columns;

            bool ColumnFiledHasSpec = false;
            foreach (string col in ColumnFields) if (Speccolumns.Contains(col)) ColumnFiledHasSpec = true;

            dt.Columns.Add("Parameter");

            //dt.Columns.Add(RowFields);
            foreach (string s in RowFields)
                dt.Columns.Add(s);

            bool SpecNeeded = false;
            foreach (var fun in Aggregate) if (fun.ToString().Contains("Spec")) SpecNeeded = true;
            if ((!ColumnFiledHasSpec) & SpecNeeded)
            {
                if (Aggregate.Contains(AggregateFunction.SpecMin)) dt.Columns.Add("SpecMin");
                if (Aggregate.Contains(AggregateFunction.SpecMax)) dt.Columns.Add("SpecMax");
                if (Aggregate.Contains(AggregateFunction.SpecTypical)) dt.Columns.Add("SpecTypical");
            }

            if (ColumnFields.Length > 0)
                foreach (string col in ColList)
                {
                    foreach (var fun in Aggregate)
                        if (!fun.ToString().Contains("Spec")) dt.Columns.Add(col + '#' + fun.ToString());  // Cretes the result columns.//
                }
            else
            {
                foreach (var fun in Aggregate)
                    if (!fun.ToString().Contains("Spec")) dt.Columns.Add(fun.ToString());  // Cretes the result columns.//
            }
            bool NoData = false;

            if (RowFields.Length > 0)
            {
                foreach (var RowName in RowList)
                {
                    DataRow row = dt.NewRow();
                    row["Parameter"] = DataField;
                    string strFilter = string.Empty;
                    string strSpecFilter = string.Empty;

                    foreach (string Field in RowFields)
                    {
                        row[Field] = RowName[Field];
                        strFilter += " and " + Field + " = '" + RowName[Field].ToString() + "'";

                        if (Speccolumns.Contains(Field)) strSpecFilter += " and " + Field + " = '" + RowName[Field].ToString() + "'";
                    }
                    strFilter = AddFilter + strFilter.Substring(5);
                    if (strSpecFilter.Length > 5) strSpecFilter = AddFilter + strSpecFilter.Substring(5);

                    if (ColumnFields.Length > 0)
                    {
                        if ((!ColumnFiledHasSpec) & SpecNeeded)
                        {
                            string filter = strSpecFilter;
                            if (Aggregate.Contains(AggregateFunction.SpecMin)) row["SpecMin"] = IncludeSpec(filter, DataField, "SpecMin");
                            if (Aggregate.Contains(AggregateFunction.SpecMax)) row["SpecMax"] = IncludeSpec(filter, DataField, "SpecMax");
                            if (Aggregate.Contains(AggregateFunction.SpecTypical)) row["SpecTypical"] = IncludeSpec(filter, DataField, "SpecTypical");
                        }

                        foreach (string col in ColList)
                        {
                            string filter = strFilter;
                            string[] strColValues = col.Split(Separator.ToCharArray(), StringSplitOptions.None);
                            for (int i = 0; i < ColumnFields.Length; i++)
                                filter += " and " + ColumnFields[i] + " = '" + strColValues[i] + "'";
                            foreach (var fun in Aggregate)
                            {
                                if (!fun.ToString().Contains("Spec"))
                                {
                                    row[col + '#' + fun.ToString()] = GetData(new string[] { filter }, DataField, fun, NSigma);
                                }
                                else
                                {
                                    //filter = strSpecFilter;
                                    //row[col + '#' + fun.ToString()] = IncludeSpec(filter, DataField, fun.ToString());
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (var fun in Aggregate)
                        {
                            string filter = strFilter;
                            if (!fun.ToString().Contains("Spec"))
                            {
                                row[fun.ToString()] = GetData(new string[] { filter }, DataField, fun, NSigma);
                                if (row[fun.ToString()].ToString() == "") { NoData = true; continue; }
                                else NoData = false;
                            }
                            else
                            {
                                filter = strSpecFilter;
                                row[fun.ToString()] = IncludeSpec(filter, DataField, fun.ToString());
                            }
                        }
                    }
                    if (!NoData) dt.Rows.Add(row);
                }
            }
            else
            {
                DataRow row = dt.NewRow();
                string strFilter = string.Empty;
                string strSpecFilter = string.Empty;

                if ((!ColumnFiledHasSpec) & SpecNeeded)
                {
                    string filter = strSpecFilter;
                    if (Aggregate.Contains(AggregateFunction.SpecMin)) row["SpecMin"] = IncludeSpec(filter, DataField, "SpecMin");
                    if (Aggregate.Contains(AggregateFunction.SpecMax)) row["SpecMax"] = IncludeSpec(filter, DataField, "SpecMax");
                    if (Aggregate.Contains(AggregateFunction.SpecTypical)) row["SpecTypical"] = IncludeSpec(filter, DataField, "SpecTypical");
                }

                foreach (string col in ColList)
                {
                    string filter = strFilter;
                    string[] strColValues = col.Split(Separator.ToCharArray(), StringSplitOptions.None);
                    for (int i = 0; i < ColumnFields.Length; i++)
                        filter += " and " + ColumnFields[i] + " = '" + strColValues[i] + "'";

                    filter = AddFilter + filter.Substring(5);
                    if (strSpecFilter.Length > 5) strSpecFilter = AddFilter + strSpecFilter.Substring(5);

                    foreach (var fun in Aggregate)
                    {
                        if (!fun.ToString().Contains("Spec")) row[col + '#' + fun.ToString()] = GetData(new string[] { filter }, DataField, fun, NSigma);
                        else
                        {
                            //filter = strSpecFilter;
                            //row[col + '#' + fun.ToString()] = IncludeSpec(filter, DataField, fun.ToString());
                        }
                    }
                }
                dt.Rows.Add(row);
            }
        }
        catch (Exception ex) { errormsg = errormsg + "," + ex.Message; }
        return dt;
    }
    public DataTable PullPivotData(string DataField, AggregateFunction[] Aggregate, string[] RowFields, string[] ColumnFields, double NSigma = 3, int[] CurrentPosition = null)
    {
        DataTable dt = new DataTable();
        DataTable dtp = new DataTable();
        try
        {
            string Separator = "$#";
            var RowList = _SourceTable.DefaultView.ToTable(true, RowFields).AsEnumerable().ToList();
            for (int index = RowFields.Count() - 1; index >= 0; index--)
                RowList = RowList.OrderBy(x => x.Field<object>(RowFields[index])).ToList();
            // Gets the list of columns .(dot) separated.
            var ColList = (from x in _SourceTable.AsEnumerable()
                           select new
                           {
                               Name = ColumnFields.Select(n => x.Field<object>(n))
                                   .Aggregate((a, b) => a += Separator + b.ToString())
                           })
                               .Distinct()
                               .OrderBy(m => m.Name);

            //dt.Columns.Add(RowFields);
            foreach (string s in RowFields)
                dt.Columns.Add(s);

            if (ColumnFields.Length > 0)
                foreach (var col in ColList)
                    foreach (var fun in Aggregate)
                        dt.Columns.Add(col.Name.ToString() + '#' + fun.ToString());  // Cretes the result columns.//
            else
                foreach (var fun in Aggregate)
                    dt.Columns.Add(fun.ToString());  // Cretes the result columns.//

            //Add Column names from Raw data
            foreach (DataColumn col in _SourceTable.Columns)
                dtp.Columns.Add(col.ColumnName.ToString());

            int rowCount = 0;
            foreach (var RowName in RowList)
            {
                DataRow row = dt.NewRow();
                rowCount++;
                string strFilter = string.Empty;

                foreach (string Field in RowFields)
                {
                    row[Field] = RowName[Field];
                    strFilter += " and " + Field + " = '" + RowName[Field].ToString() + "'";
                }
                strFilter = strFilter.Substring(5);

                if (ColumnFields.Length > 0)
                    foreach (var col in ColList)
                    {
                        string filter = strFilter;
                        string[] strColValues = col.Name.ToString().Split(Separator.ToCharArray(), StringSplitOptions.None);
                        for (int i = 0; i < ColumnFields.Length; i++)
                            filter += " and " + ColumnFields[i] + " = '" + strColValues[i] + "'";
                        if (rowCount == CurrentPosition[0] + 1)
                        {
                            DataRow[] FilteredRows = _SourceTable.Select(filter);
                            foreach (DataRow rowEntry in FilteredRows)
                            {
                                dtp.ImportRow(rowEntry);
                            }
                        }
                    }
                else
                {
                    string filter = strFilter;
                    if (rowCount == CurrentPosition[0] + 1)
                    {
                        DataRow[] FilteredRows = _SourceTable.Select(filter);
                        foreach (DataRow rowEntry in FilteredRows)
                        {
                            dtp.ImportRow(rowEntry);
                        }
                    }
                }
                //dt.Rows.Add(row);
            }
        }
        catch (Exception ex) { }
        return dtp;
    }
    /// <summary>
    /// Retrives the data for matching RowField value and ColumnFields values with Aggregate function applied on them.
    /// </summary>
    /// <param name="Filter">DataTable Filter condition as a string</param>
    /// <param name="DataField">The column name which needs to spread out in Data Part of the Pivoted table</param>
    /// <param name="Aggregate">Enumeration to determine which function to apply to aggregate the data</param>
    /// <returns></returns>
    private object IncludeSpec(string Filter, string DataField, string Type)
    {
        if (Type == "SpecMin") Type = "_Min";
        if (Type == "SpecMax") Type = "_Max";
        if (Type == "SpecTypical") Type = "_Typical";

        try
        {
            DataRow[] FilteredRows = _SpecTable.Select(Filter);
            object[] objList = FilteredRows.Select(x => x.Field<object>(DataField + Type)).ToArray();

            return objList.Count() == 0 ? null : objList.Count() > 1 ? GetMin(objList) : objList[0].ToString();
            //Todo Return error when the count is more than one, currently returning Min
        }
        catch (Exception ex)
        {
            return "No Spec";
        }
    }
    public void FilterSource(string[] Headers, string[] entries)
    {
        DataTable dtp = new DataTable();
        var RowList = _SourceTable.DefaultView.ToTable(true, Headers).AsEnumerable().ToList();

        foreach (DataColumn col in _SourceTable.Columns)
            dtp.Columns.Add(col.ColumnName.ToString());

        string strFilter = string.Empty;
        foreach (string Field in Headers)
        {
            foreach (string entry in entries)
            {
                strFilter += " or " + Field + " = '" + entry + "'";
            }
        }
        if (strFilter.Length < 5) return;
        strFilter = strFilter.Substring(4);

        DataRow[] FilteredRows = _SourceTable.Select(strFilter);
        foreach (DataRow rowEntry in FilteredRows)
        {
            dtp.ImportRow(rowEntry);
        }
        _SourceTable = dtp;
    }
    /// <summary>
    /// Retrives the data for matching RowField value and ColumnFields values with Aggregate function applied on them.
    /// </summary>
    /// <param name="Filter">DataTable Filter condition as a string</param>
    /// <param name="DataField">The column name which needs to spread out in Data Part of the Pivoted table</param>
    /// <param name="Aggregate">Enumeration to determine which function to apply to aggregate the data</param>
    /// <returns></returns>
    public object GetData(string[] Filters, string DataField, AggregateFunction Aggregate, double NSigma = 3)
    {
        try
        {
            List<DataRow[]> FilteredRows = new List<DataRow[]>();
            foreach (string filter in Filters)
            {
                FilteredRows.Add(_SourceTable.Select(filter));
            }
            object[] objList = FilteredRows.Select(x => x.Select(y => y.Field<object>(DataField))).SelectMany(i => i).ToArray();
            switch (Aggregate)
            {
                case AggregateFunction.Average:
                    return GetAverage(objList);
                case AggregateFunction.Count:
                    return objList.Count();
                case AggregateFunction.Exists:
                    return (objList.Count() == 0) ? "False" : "True";
                case AggregateFunction.First:
                    return GetFirst(objList);
                case AggregateFunction.Last:
                    return GetLast(objList);
                case AggregateFunction.Max:
                    return GetMax(objList);
                case AggregateFunction.Min:
                    return GetMin(objList);
                case AggregateFunction.Sum:
                    return GetSum(objList);
                case AggregateFunction.Stdev:
                    return GetStdev(objList);
                case AggregateFunction.NSigmaPos:
                    return GetNSigmaPos(objList, NSigma);
                case AggregateFunction.NSigmaNeg:
                    return GetNSigmaNeg(objList, NSigma);
                case AggregateFunction.FullArray:
                    return GetFullArray(objList);
                case AggregateFunction.NSigmaNegMin:
                    List<object> minimumNSigmaNeg = new List<object>();
                    foreach (var e in FilteredRows)
                    {
                        minimumNSigmaNeg.Add(GetNSigmaNeg(e.Select(y => y.Field<object>(DataField)).ToArray(), NSigma));
                    }
                    return GetMin(minimumNSigmaNeg.ToArray());
                case AggregateFunction.NSigmaPosMax:
                    List<object> maximumNSigmaPos = new List<object>();
                    foreach (var e in FilteredRows)
                    {
                        maximumNSigmaPos.Add(GetNSigmaPos(e.Select(y => y.Field<object>(DataField)).ToArray(), NSigma));
                    }
                    return GetMax(maximumNSigmaPos.ToArray());
                case AggregateFunction.NSigmaNegCorrespondingStdDev:
                    List<Tuple<object, object>> correspondingNSigmaNeg = new List<Tuple<object, object>>();
                    foreach (var e in FilteredRows)
                    {
                        Tuple<object, object> stdDevNSigmaNed = new Tuple<object, object>(GetStdev(e.Select(y => y.Field<object>(DataField)).ToArray()), GetNSigmaNeg(e.Select(y => y.Field<object>(DataField)).ToArray(), NSigma));
                        correspondingNSigmaNeg.Add(stdDevNSigmaNed);
                    }
                    return correspondingNSigmaNeg.OrderBy(x => x.Item2).ToList()[0].Item1;
                case AggregateFunction.NSigmaPosCorrespondingStdDev:
                    List<Tuple<object, object>> correspondingPSigmaPos = new List<Tuple<object, object>>();
                    foreach (var e in FilteredRows)
                    {
                        Tuple<object, object> stdDevNSigmaPos = new Tuple<object, object>(GetStdev(e.Select(y => y.Field<object>(DataField)).ToArray()), GetNSigmaPos(e.Select(y => y.Field<object>(DataField)).ToArray(), NSigma));
                        correspondingPSigmaPos.Add(stdDevNSigmaPos);
                    }
                    return correspondingPSigmaPos.OrderByDescending(x => x.Item2).ToList()[0].Item1;
                default:
                    return null;
            }
        }
        catch (Exception ex)
        {
            return "#Error";
        }
    }
    public string GetFullArray(object[] objList)
    {
        string[] strList = Array.ConvertAll(objList, p => (p ?? String.Empty).ToString());
        string strbuild = string.Empty;
        foreach (string str in strList) strbuild = strbuild + str + ",";
        return strbuild;
    }
    public object GetAverage(object[] objList)
    {
        return objList.Count() == 0 ? null : (object)(Convert.ToDecimal(GetSum(objList)) / objList.Count());
    }
    public object GetAverageOmitNull(object[] objList)
    {
        object[] objListNullFree = FilterNull(objList);
        return objList.Count() == 0 ? null : (object)(Convert.ToDecimal(GetSum(objListNullFree)) / objListNullFree.Count());
    }
    public object GetNSigmaNeg(object[] objList, double Nsigma)
    {
        return objList.Count() == 0 ? null : (object)(Convert.ToDouble(GetAverage(objList)) - (Nsigma * (Convert.ToDouble(GetStdev(objList)))));
    }
    public object GetNSigmaPos(object[] objList, double Nsigma)
    {
        return objList.Count() == 0 ? null : (object)(Convert.ToDouble(GetAverage(objList)) + (Nsigma * (Convert.ToDouble(GetStdev(objList)))));
    }
    public object[] FilterNull(object[] objList)
    {
        objList = objList.Where(x => x != null).ToArray();
        objList = objList.Where(x => x.ToString() != "NaN").ToArray();
        return objList;
    }
    public double[] ConvertDouble(object[] objList)
    {
        return Array.ConvertAll<object, double>(objList, o => Convert.ToDouble(o));
    }
    public object GetStdev(object[] objList)
    {
        double ret = 0;
        if (objList.Count() > 0)
        {
            //Compute the Average      
            double avg = (Convert.ToDouble(GetAverage(objList)));
            //Perform the Sum of (value-avg)_2_2      
            double sum = objList.Sum(d => Math.Pow(Convert.ToDouble(d) - avg, 2));
            //Put it all together      
            ret = Math.Sqrt((sum) / (objList.Count() - 1));
        }
        return objList.Count() == 0 ? null : (object)ret;
    }
    public object GetSum(object[] objList)
    {
        return objList.Count() == 0 ? null : (object)(objList.Aggregate(new decimal(), (x, y) => x += Convert.ToDecimal(y)));
    }
    public object GetFirst(object[] objList)
    {
        return (objList.Count() == 0) ? null : objList.First();
    }
    public object GetLast(object[] objList)
    {
        return (objList.Count() == 0) ? null : objList.Last();
    }
    public object GetMax(object[] objList)
    {
        return (objList.Count() == 0) ? null : (object)ConvertDouble(FilterNull(objList)).Max();
    }
    public object GetMin(object[] objList)
    {
        return (objList.Count() == 0) ? null : (object)ConvertDouble(FilterNull(objList)).Min();
    }
}

public enum AggregateFunction
{
    Count = 1,
    Sum = 2,
    First = 3,
    Last = 4,
    Average = 5,
    Max = 6,
    Min = 7,
    Exists = 8,
    Stdev = 9,
    NSigmaNeg = 10,
    NSigmaPos = 11,
    SpecMin = 12,
    SpecMax = 13,
    SpecTypical = 14,
    FullArray = 15,
    NSigmaNegMin = 16,
    NSigmaPosMax = 17,
    NSigmaNegCorrespondingStdDev = 18,
    NSigmaPosCorrespondingStdDev = 19
}
