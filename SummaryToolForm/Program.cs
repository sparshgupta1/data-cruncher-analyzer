using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SummaryToolForm
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            { 
                Application.Run(new SummaryToolFrm()); 
            }
            catch
            {
                //Delete the last loaded file
                string path = Path.Combine(Path.GetTempPath() + "SummaryTool.Settings.xml");
                if (System.IO.File.Exists(path)) System.IO.File.Delete(path);
                Application.Run(new SummaryToolFrm());
            }

        }
    }
}
