﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.IO;

namespace SummaryTool
{
    class ExportToCsv
    {
        public static void ToCsV(DataGridView dGV, string filename, bool append = false)
        {

            string stOutput = "";
            // Export titles:
            string sHeaders = "";

            for (int j = 0; j < dGV.Columns.Count; j++)
                sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText).Replace("null","") + ",";
            //stOutput += sHeaders + "\r\n";
            // Export data.
            for (int i = 0; i < dGV.RowCount - 1; i++)
            {
                string stLine = "";
                for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
                    stLine = stLine.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + ",";
                stOutput += stLine + "\r\n";
            }
            Encoding utf16 = Encoding.GetEncoding(1254);
            if (!File.Exists(filename)) stOutput = sHeaders + "\r\n" + stOutput;
            byte[] output = utf16.GetBytes(stOutput);

            FileStream fs;
            if (append == true)
            {
                fs = new FileStream(filename, FileMode.Append);
            }
            else
            {
                fs = new FileStream(filename, FileMode.Create);
            }

            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(output, 0, output.Length); //write the encoded file
            bw.Flush();
            bw.Close();
            fs.Close();
        }
    }
}
