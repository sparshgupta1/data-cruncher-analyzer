using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using XmlSerialization;
using System.IO;
namespace SummaryTool
{
    public class SummaryMainClass
    {
        public Settings inputSetting; // = new settings();
        static object xobj = new Settings();
        public Serializer sl = new Serializer(ref xobj);
        public DataTable dt;
        public DataTable sdt;
        public Pivot pvt;
        public string errormsg;
        public static double NSigma;

        public SummaryMainClass()
        {

            inputSetting = (Settings)xobj;
            string msg = null;
            if (inputSetting != null)
            {
                if (!System.IO.File.Exists(inputSetting.DataPath)) msg = msg + "Data Path is invalid!";
                if (!System.IO.File.Exists(inputSetting.SpecPath)) msg = msg + "Spec Path is invalid!";
                if (!(msg == null)) sl.showForm(msg);
                if (System.IO.File.Exists(inputSetting.DataPath))
                {
                    dt = new DataTable();
                    sdt = new DataTable();
                    dt = ExcelLayer.GetDataTable(inputSetting.DataPath);
                    sdt = ExcelLayer.GetDataTable(inputSetting.SpecPath);

                    if (dt != null) ChangeDataTable(ref dt);
                    if (sdt != null) ChangeDataTable(ref sdt);

                    pvt = new Pivot(dt, sdt);
                }
            }
        }
        public void showInput()
        {
            sl.showForm();
        }
        public DataTable RawData()
        {
            return dt;
        }
        public DataTable HighSummaryData(string DataString = null,string[] Iterator = null, string[] RowFields = null, string[] ColumnFields = null, AggregateFunction[] Aggregate = null, double NSigma1 = 3)
        {
            var errors = new List<string>();
            NSigma = NSigma1;          
            DataTable sdt = pvt.ManagedPivotData(DataString, Aggregate,Iterator, RowFields, ColumnFields, NSigma1);
            errormsg = pvt.errormsg;
            RestoreDataTable(ref sdt);

            return sdt;
        }

        public DataTable SummaryData(string DataString = null, string[] RowFields = null, string[] ColumnFields = null, AggregateFunction[] Aggregate = null, double NSigma1 = 3)
        {
            var errors = new List<string>();
            NSigma = NSigma1;
            //Sample usage as below
            //DataTable sdt = pvt.PivotData("RiseTimeFUNCTION1", AggregateFunction.Max, new string[] { "SELECTLINES", "PIN_NAME" }, new string[] { "Voltage1" });
            DataTable sdt = pvt.PivotData(DataString, Aggregate, RowFields, ColumnFields, NSigma1);
            errormsg = pvt.errormsg;
            RestoreDataTable(ref sdt);

            return sdt;
        }
        public DataTable PullData(string DataString = null, string[] RowFields = null, string[] ColumnFields = null, AggregateFunction[] Aggregate = null, double NSigma1 = 3, int[] CurrentPosition = null)
        {
            var errors = new List<string>();
            NSigma = NSigma1;

            DataTable sdt = pvt.PullPivotData(DataString, Aggregate, RowFields, ColumnFields, NSigma1, CurrentPosition);
            RestoreDataTable(ref sdt);
            //pvt.printDataTable(sdt);
            return sdt;
        }
        public static string RestoreData(string str)
        {
            str = str.Replace("NSigmaPos", "+" + NSigma.ToString() + "Sigma");
            str = str.Replace("NSigmaNeg", "-" + NSigma.ToString() + "Sigma");
            return str.Replace("qspaceq", " ");
        }
        public static string ChangeData(string str)
        {
            return str.Replace(" ", "qspaceq");
        }

        public static void RestoreDataTable(ref DataTable dt)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dt.Columns[i].ColumnName = RestoreData(dt.Columns[i].ColumnName);
            }
        }
        public static void ChangeDataTable(ref DataTable dt)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dt.Columns[i].ColumnName = ChangeData(dt.Columns[i].ColumnName);
            }

        }
    }
    public class Settings
    {

        //public string SpecPath = "\\\\flrblr001\\windows_data\\From-SLI-Fileserver\\E\\Product Engineering\\Sundar\\svn\\trunk\\CommonBase\\DeviceSpecific\\Si53300\\sample_data\\Spec.csv";
        public string SpecPath = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory()))), "Spec.csv");
        //public string DataPath = "\\\\flrblr001\\windows_data\\From-SLI-Fileserver\\E\\Product Engineering\\Sundar\\svn\\trunk\\CommonBase\\DeviceSpecific\\Si53300\\sample_data\\LVPECLmini.csv";
        public string DataPath = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory()))), "Data.csv");
        public string DataSeperator = ",";
    }
}
