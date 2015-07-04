using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Excel; //use the reference in your code

/// <summary>
/// Acts as a DataBase Layer
/// </summary>
public class ExcelLayer
{
    public static List<string> errors = new List<string>();
    public ExcelLayer()
    {
        //
        // TODO: Add constructor logic here
        //

        //For more complex data refer http://www.codeproject.com/Articles/11698/A-Portable-and-Efficient-Generic-Parser-for-Flat-F
        errors = new List<string>();
    }
    /// <summary>
    /// Retireves the data from Excel Sheet to a System.Data.DataTable.
    /// </summary>
    /// <param name="FileName">File Name along with path from the root folder.</param>
    /// <param name="TableName">Name of the Table of the Excel Sheet. Sheet1$ if no table.</param>
    /// <returns></returns>
    public static System.Data.DataTable GetDataTable(string strPath , string separatorChar = ",", bool isFirstRowHeader = true)
    {
        //var errors = new List<string>();
        if (!File.Exists(strPath)) errors.Add("File " + strPath + "doesn't exist!");
        string SheetName = "";
        string FileExtension = "";
        try
        {
            //File extension
            FileExtension = Path.GetExtension(strPath).ToLower();
            SheetName = Path.GetFileNameWithoutExtension(strPath);

            string tempFile = null;
            if (IsFileLocked(strPath))
            {
                string filename = Path.GetFileName(strPath);
                string path = Path.GetDirectoryName(strPath);
                tempFile = path + "\\tempfile_" + filename;
                if (System.IO.File.Exists(tempFile)) System.IO.File.Delete(tempFile);
                System.IO.File.Copy(strPath, tempFile);
                strPath = tempFile;
            }

            System.Data.DataTable dt = null;

            //csv file parser
            if (FileExtension.Contains("csv")) dt = CsvFileToDatatable(strPath, true);

            //xls file parse
            if (FileExtension.Contains("xls")) dt = ExcelToDatatable(strPath, SheetName, isFirstRowHeader);

            //Other file type parse
            if (!FileExtension.Contains("csv") & !FileExtension.Contains("xls")) dt = TextFile2DataTable(strPath, separatorChar);

            //Clear the temp file if any
            if (tempFile != null) if (System.IO.File.Exists(tempFile)) System.IO.File.Delete(tempFile);
            return dt;
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(ex.Message);
            //Log your exception here.//
        }
        return (System.Data.DataTable)null;
    }
    /// <suninary>
    /// Th/s function is used to check specified file being used or not
    /// </sunrary>
    /// <param name=’file’>Filelnfo of required file</param>
    /// <returns>If that specified file is being processed
    /// or not found is return true</returns>
    public static Boolean IsFileLocked(string path)
    {
        FileInfo file = new FileInfo(path);
        System.IO.FileStream stream = null;
        try
        {
            //Dont change FileAccess to ReadWrite,
            //because if a file is in readOnly, it fails.
            stream = file.Open
            (
            FileMode.Open,
            FileAccess.Read,
            FileShare.None
            );
        }
        catch (IOException)
        {
            //the file is unavailable because it is:
            //still being written to
            //or being processed by another thread
            //or does not exist (has already been processed)
            return true;
        }
        finally
        {
            if (stream != null)
                stream.Close();
        }
        //file is not locked
        return false;
    }

    public static System.Data.DataTable ExcelToDatatable(string strPath, string TableName, bool IsFirstRowHeader)
    {
        DataSet ds = new DataSet();
        //String sConnectionString = "";
        //Todo: get the excel sheet name from sheet number

        //sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=" + strPath + "; " + "Extended Properties=Excel 8.0;";
        //OleDbConnection objConn = new OleDbConnection(sConnectionString);
        //objConn.Open();
        //OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM [" + TableName + "$" + "] ", objConn);
        //OleDbDataAdapter objAdapter1 = new OleDbDataAdapter();
        //objAdapter1.SelectCommand = objCmdSelect;
        //objAdapter1.Fill(ds);
        //objConn.Close();
        //return ds.Tables[0];


        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strPath + ";Extended Properties='Excel 12.0 xml;HDR=YES;'"); //12 to 14
        //OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.12.0;Data Source=" + strPath + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");
        try
        {
            connection.Open();
            System.Data.DataTable Sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            foreach (DataRow dr in Sheets.Rows)
            {
                string sht = dr[2].ToString().Replace("'", "");
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter("select * from [" + sht + "]", connection);
                dataAdapter.Fill(ds);
                connection.Close();
                return ds.Tables[0];

            }
        }
        catch 
        {
        }
            return ds.Tables[0];

    }
    public static System.Data.DataTable CsvFileToDatatable(string path, bool IsFirstRowHeader)
    {

        //Check for incompatibility and fix it
        //using (StreamReader reader = new StreamReader(path))
        //{
        //    string line1 = reader.ReadLine();
        //    var headers = line1.Split(',');
        //    string newHeader = "";
        //    var data = new System.Data.DataTable();
        //    int repeatCount = 1;
        //    foreach (var hdr in headers)
        //    {
        //        if (hdr == "") { newHeader = newHeader + "null"; }
        //        if (!data.Columns.Contains(hdr)) { data.Columns.Add(hdr); newHeader = newHeader + "," + hdr; }
        //        else
        //        {
        //            //Here goes repeated data
        //            data.Columns.Add(hdr + "repeated" + repeatCount.ToString());
        //            repeatCount++;
        //        }
        //    }
        //}

        string header = "No";
        string sql = string.Empty;
        System.Data.DataTable dataTable = null;
        string pathOnly = string.Empty;
        string fileName = string.Empty;
        try
        {
            pathOnly = Path.GetDirectoryName(path);
            fileName = Path.GetFileName(path);
            sql = @"SELECT * FROM [" + fileName + "]";
            if (IsFirstRowHeader)
            {
                header = "Yes";
            }
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +
            ";Extended Properties=\"Text;HDR=" + header + ";FMT=Delimited;IMEX=1\""))
            {
                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        dataTable = new System.Data.DataTable();
                        dataTable.Locale = System.Globalization.CultureInfo.CurrentCulture;
                        adapter.Fill(dataTable);
                    }
                }
            }
        }
        finally
        {
        }
        return dataTable;
    }
    public static System.Data.DataTable TextFile2DataTable(string filename, string separatorChar)
    {
        var errors = new List<string>();
        var data = new System.Data.DataTable();
        int repeatCount = 1;
        int NullColumnCount = 1;
        try
        {
            var reader = ReadAsLines(filename);


            //this assume the first record is filled with the column names
            var headers = reader.First().Split(',');
            foreach (var header in headers)
            {
                if (header == "") { data.Columns.Add("null" + NullColumnCount); NullColumnCount++; }
                if (!data.Columns.Contains(header)) data.Columns.Add(header);
                else
                {
                    //Here goes repeated data
                    data.Columns.Add(header + "repeated" + repeatCount.ToString());
                    repeatCount++;
                }
            }

            var records = reader.Skip(1);
            foreach (var record in records)
            {
                data.Rows.Add(record.Split(','));
            }
        }
        catch (Exception ex)
        {
            errors.Add(ex.Message);
        }
        return data;
    }

    static IEnumerable<string> ReadAsLines(string filename)
    {
        using (StreamReader reader = new StreamReader(filename))
            while (!reader.EndOfStream)
                yield return reader.ReadLine();
    }

    public static System.Data.DataTable TextFile2DataTable2(string filename, string separatorChar)
    {
        var errors = new List<string>();
        var table = new System.Data.DataTable("StringLocalization");
        using (var sr = new StreamReader(filename, System.Text.Encoding.Default))
        {
            string line;
            var i = 0;
            while (sr.Peek() >= 0)
            {
                try
                {
                    line = sr.ReadLine();
                    if (string.IsNullOrEmpty(line)) continue;
                    var values = line.Split(new[] { separatorChar }, StringSplitOptions.None);
                    var row = table.NewRow();
                    for (var colNum = 0; colNum < values.Length; colNum++)
                    {
                        var value = values[colNum];
                        if (i == 0)
                        {
                            table.Columns.Add(value, typeof(String));
                        }
                        else
                        {
                            row[table.Columns[colNum]] = value;
                        }
                    }
                    if (i != 0) table.Rows.Add(row);
                }
                catch (Exception ex)
                {
                    errors.Add(ex.Message);
                }
                i++;
            }
        }
        return table;
    }
}
