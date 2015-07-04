using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SummaryTool;

namespace SummaryToolForm
{
    public partial class Form1Old : Form
    {
        public Form1Old()
        {
            InitializeComponent();
            summaryInit();
        }
        public void summaryInit()
        {
            //*** cb***
            //Refer C:\Users\suandich\Downloads\DragAndDropListView_demo\Form1.cs
            SummaryMainClass sm = new SummaryMainClass();
            DataTable dtRaw = sm.RawData();
            for (int i = 0; i < dtRaw.Columns.Count; i++)
            {
                System.Diagnostics.Debug.Write(sm.dt.Columns[i].ToString() + " | ");
                listView1.Items.Add(sm.dt.Columns[i].ToString());
            }

            string[] Rowfield = new string[listView3.Items.Count];
            listView3.Items.CopyTo(Rowfield, 0);

            string[] Columnfield = new string[listView2.Items.Count];
            listView2.Items.CopyTo(Columnfield, 0);

            string[] Datafield = new string[listView4.Items.Count];
            listView4.Items.CopyTo(Datafield, 0);

            System.Collections.ArrayList list = new System.Collections.ArrayList();
            list.Add("Cat");
            list.Add("Zebra");
            list.Add("Dog");
            list.Add("Cow");

            DataTable dtSum = sm.SummaryData("",Rowfield,Columnfield);
            dataGridView1.DataSource = dtSum.DefaultView;
        }
    }
}
