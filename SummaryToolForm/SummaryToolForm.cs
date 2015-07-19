using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using SummaryTool;
using FormSerializor;
using System.Collections.Generic;
using System.Windows.Forms.DataVisualization.Charting;
using Update;
using System.Linq;
using System.IO;

namespace SummaryToolForm
{
    /// <summary>
    /// Summary description for Form3.
    /// </summary>
    public class SummaryToolFrm : System.Windows.Forms.Form
    {
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private DragNDrop.DragNDropListView XaxisRowDNDLV;
        private DragNDrop.DragNDropListView YAxisOutputDNDLV;
        private DragNDrop.DragNDropListView InputRowDNDLV;
        private ColumnHeader columnHeader5;
        private ColumnHeader columnHeader6;
        private DragNDrop.DragNDropListView AllEntryDNDLV;
        private ColumnHeader columnHeader7;
        private ColumnHeader columnHeader8;
        private DataGridView SumdataGridView;
        private IContainer components;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem menuToolStripMenuItem;
        private ToolStripMenuItem selectInputsToolStripMenuItem;
        private ToolStripMenuItem restoreToolStripMenuItem;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel toolStripStatusLabel1;
        private ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private SplitContainer splitContainer1;
        private SplitContainer splitContainer2;
        private SplitContainer splitContainer3;
        private SplitContainer splitContainer4;
        private SplitContainer splitContainer5;
        private SplitContainer splitContainer6;
        private TextBox NSigmatextBox;
        private Label label4;
        private DragNDrop.DragNDropListView StatisticsSelectedDNDLV;
        private ColumnHeader columnHeader11;
        private ColumnHeader columnHeader12;
        private DragNDrop.DragNDropListView StatisticsDNDLV;
        private ColumnHeader columnHeader9;
        private ColumnHeader columnHeader10;
        private Label label5;
        private SplitContainer splitContainer7;
        private DataGridView FilterdataGridView;

        private ToolStripMenuItem exportToCSVToolStripMenuItem;
        private ToolStripMenuItem table1ToolStripMenuItem;
        private ToolStripMenuItem table2ToolStripMenuItem;
        private SplitContainer splitContainer8;
        private DataGridView BpdataGridView;
        private MenuStrip menuStrip2;
        private ToolStripMenuItem summaryTableToolStripMenuItem;
        private MenuStrip menuStrip3;
        private ToolStripMenuItem filterDataToolStripMenuItem;
        private MenuStrip menuStrip4;
        private ToolStripMenuItem chartSummaryToolStripMenuItem;
        private ToolStripMenuItem disableSummaryToolStripMenuItem;
        private ToolStripMenuItem disableChartToolStripMenuItem;
        private ToolStripMenuItem disableChartSummaryToolStripMenuItem;

        public DataTable dtSum;
        public SummaryMainClass sm;
        public string[] Rowfield;
        public string[] Columnfield;
        public string[] Datafield;
        public AggregateFunction[] Aggregate;
        private TextBox textBox1;
        private Label label6;
        private TabControl tabControl1;
        private TabPage tabPageEdit;
        private TabPage tabPageSummaryTool;
        private ToolStripMenuItem createSpecToolStripMenuItem;
        private ToolStripMenuItem exitToolStripMenuItem;
        private ToolStripMenuItem chartSummaryToolStripMenuItem1;
        public double NSigma = 3;
        //public DataCrunch.DataCrunchClass dc;
        private ToolStripMenuItem refreshToolStripMenuItem;
        private ToolStripMenuItem releaseToQAToolStripMenuItem;
        private ToolStripMenuItem releaseProductionToolStripMenuItem;

        public DtPlots DtPlots1;
        private ToolStripMenuItem AdminMode;
        private ToolStripMenuItem updateToolStripMenuItem;
        private DataGridView EditdataGridView;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripMenuItem disableSummaryToolStripMenuItem1;
        private ToolStripMenuItem disableChartDCToolStripMenuItem;
        private ToolStripMenuItem disableChartSummaryDUToolStripMenuItem;
        private ToolStripMenuItem disableInputSectionToolStripMenuItem;
        private ToolStripMenuItem disableInputSectionToolStripMenuItem1;
        private NotifyIcon notifyIcon1;
        private ToolStripMenuItem minimizeToTrayToolStripMenuItem;
        private ToolStripMenuItem exportSummaryToPPTToolStripMenuItem;
        private ToolStripMenuItem chartSummaryPCToolStripMenuItem;
        private ToolStripMenuItem summaryPSToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator2;
        private MenuStrip menuStrip8;
        private MenuStrip menuStrip9;
        private ToolStripMenuItem statisticsToolStripMenuItem;
        private ToolStripMenuItem specificationNToolStripMenuItem;
        private SplitContainer splitContainer9;
        private DragNDrop.DragNDropListView InputColumnDNDLV;
        private ColumnHeader columnHeader13;
        private ColumnHeader columnHeader14;
        private MenuStrip menuStrip6;
        private ToolStripMenuItem inputRowToolStripMenuItem;
        private MenuStrip menuStrip10;
        private ToolStripMenuItem inputColumnToolStripMenuItem;
        private SplitContainer splitContainer11;
        private DragNDrop.DragNDropListView ChartSeriesRowDNDLV;
        private ColumnHeader columnHeader17;
        private ColumnHeader columnHeader18;
        private SplitContainer splitContainer10;
        private MenuStrip menuStrip5;
        private ToolStripMenuItem yAxisToolStripMenuItem;
        private DragNDrop.DragNDropListView YAxisCombinedOutputDNDLV;
        private ColumnHeader columnHeader15;
        private ColumnHeader columnHeader16;
        private MenuStrip menuStrip11;
        private ToolStripMenuItem multiOutputToolStripMenuItem;
        private SplitContainer splitContainer12;
        private DragNDrop.DragNDropListView XaxisColumnDNDLV;
        private ColumnHeader columnHeader19;
        private ColumnHeader columnHeader20;
        private SplitContainer splitContainer13;
        private DragNDrop.DragNDropListView ChartSeriesColumnDNDLV;
        private ColumnHeader columnHeader21;
        private ColumnHeader columnHeader22;
        private MenuStrip menuStrip7;
        private ToolStripMenuItem chartSeriesRowToolStripMenuItem;
        private MenuStrip menuStrip13;
        private ToolStripMenuItem chartSeriesColumnToolStripMenuItem;
        private MenuStrip menuStrip12;
        private ToolStripMenuItem xAxisRowToolStripMenuItem;
        private MenuStrip menuStrip14;
        private ToolStripMenuItem xAxisColumnToolStripMenuItem;
        private ToolStripMenuItem revisionToolStripMenuItem;

        public InputFields InputFields1 = new InputFields();
        private SplitContainer splitContainer14;
        private DataGridView HighSummaryDataGridView;
        private ToolStripMenuItem highLevelSummaryToolStripMenuItem;

        PopupForm PopupForm1 = new PopupForm();

        public SummaryToolFrm()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //

            SubscribePopups();
            InitializeInputFields();
            DtPlots1 = new DtPlots(InputFields1, new DataGridView[] { SumdataGridView, HighSummaryDataGridView, FilterdataGridView, BpdataGridView, EditdataGridView }, chart1);

            DtPlots1.summaryInit();
            FormSerializor.FormSerilizor.DeSerialise(this);

            //Check the minimum inputs needed to plot/summary 
            if (InputRowDNDLV.Items.Count == 0 & AllEntryDNDLV.Items.Count == 0 & XaxisRowDNDLV.Items.Count == 0 & YAxisOutputDNDLV.Items.Count == 0 & InputRowDNDLV.Items.Count == 0 &
                 InputColumnDNDLV.Items.Count == 0 & YAxisOutputDNDLV.Items.Count == 0 & YAxisCombinedOutputDNDLV.Items.Count == 0 & XaxisColumnDNDLV.Items.Count == 0)
            {
                SummaryPlotInit();
            }

            AdminForminit();
            try
            { refreshDgv(); }
            catch { InitializeInputFields(); }
        }
        public void SummaryToolFrmFresh()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
            SubscribePopups();
            InitializeInputFields();
            DtPlots1 = new DtPlots(InputFields1, new DataGridView[] { SumdataGridView, HighSummaryDataGridView, FilterdataGridView, BpdataGridView, EditdataGridView }, chart1);
        }
        private void SubscribePopups()
        {
        }
        public bool InitializeInputFields()
        {
            //Clear the chart
            this.chart1.Legends.Clear();
            chart1.Series.Clear();
            chart1.Titles.Clear();

            //Todo: Clear the chart after every refresh if (chart1.ChartAreas.Count > 0)  chart1.ChartAreas.Clear();

            //Get all rows
            InputFields1.SumRowfield = new string[InputRowDNDLV.Items.Count + XaxisRowDNDLV.Items.Count + ChartSeriesRowDNDLV.Items.Count];

            for (int i = 0; i < InputRowDNDLV.Items.Count; i++)
                InputFields1.SumRowfield[i] = SummaryMainClass.ChangeData(InputRowDNDLV.Items[i].Text);

            for (int i = 0; i < XaxisRowDNDLV.Items.Count; i++)
                InputFields1.SumRowfield[i + InputRowDNDLV.Items.Count] = SummaryMainClass.ChangeData(XaxisRowDNDLV.Items[i].Text);

            for (int i = 0; i < ChartSeriesRowDNDLV.Items.Count; i++)
                InputFields1.SumRowfield[i + InputRowDNDLV.Items.Count + XaxisRowDNDLV.Items.Count] = SummaryMainClass.ChangeData(ChartSeriesRowDNDLV.Items[i].Text);

            //Get all columns
            InputFields1.SumColumnfield = new string[InputColumnDNDLV.Items.Count + XaxisColumnDNDLV.Items.Count + ChartSeriesColumnDNDLV.Items.Count];

            for (int i = 0; i < InputColumnDNDLV.Items.Count; i++)
                InputFields1.SumColumnfield[i] = SummaryMainClass.ChangeData(InputColumnDNDLV.Items[i].Text);

            for (int i = 0; i < XaxisColumnDNDLV.Items.Count; i++)
                InputFields1.SumColumnfield[i + InputColumnDNDLV.Items.Count] = SummaryMainClass.ChangeData(XaxisColumnDNDLV.Items[i].Text);

            for (int i = 0; i < ChartSeriesColumnDNDLV.Items.Count; i++)
                InputFields1.SumColumnfield[i + InputColumnDNDLV.Items.Count + XaxisColumnDNDLV.Items.Count] = SummaryMainClass.ChangeData(ChartSeriesColumnDNDLV.Items[i].Text);

            if ((InputFields1.SumRowfield.Length < 1 && InputFields1.SumColumnfield.Length < 1)) { toolStripStatusLabel1.Text = "Please Select Input data!"; return false; }

            //If there are no inputs or outputs then skip
            if (YAxisOutputDNDLV.Items.Count < 1 && YAxisCombinedOutputDNDLV.Items.Count < 1) { toolStripStatusLabel1.Text = "Please Select Output data!"; return false; }

            //Get PlotSeries
            InputFields1.PlotSeries = new string[ChartSeriesRowDNDLV.Items.Count + ChartSeriesColumnDNDLV.Items.Count];

            for (int i = 0; i < ChartSeriesRowDNDLV.Items.Count; i++)
                InputFields1.PlotSeries[i] = SummaryMainClass.ChangeData(ChartSeriesRowDNDLV.Items[i].Text);

            for (int i = 0; i < ChartSeriesColumnDNDLV.Items.Count; i++)
                InputFields1.PlotSeries[i + ChartSeriesRowDNDLV.Items.Count] = SummaryMainClass.ChangeData(ChartSeriesColumnDNDLV.Items[i].Text);

            //Get X -axis 
            InputFields1.Plotx = new string[XaxisRowDNDLV.Items.Count + XaxisColumnDNDLV.Items.Count];

            for (int i = 0; i < XaxisRowDNDLV.Items.Count; i++)
                InputFields1.Plotx[i] = SummaryMainClass.ChangeData(XaxisRowDNDLV.Items[i].Text);

            for (int i = 0; i < XaxisColumnDNDLV.Items.Count; i++)
                InputFields1.Plotx[i + XaxisRowDNDLV.Items.Count] = SummaryMainClass.ChangeData(XaxisColumnDNDLV.Items[i].Text);

            //Get Y - Axis
            InputFields1.Ploty = new string[YAxisOutputDNDLV.Items.Count];

            for (int i = 0; i < YAxisOutputDNDLV.Items.Count; i++)
                InputFields1.Ploty[i] = SummaryMainClass.ChangeData(YAxisOutputDNDLV.Items[i].Text);

            //Get Y - Axis Combined
            InputFields1.PlotyCombined = new string[YAxisCombinedOutputDNDLV.Items.Count];

            for (int i = 0; i < YAxisCombinedOutputDNDLV.Items.Count; i++)
                InputFields1.PlotyCombined[i] = SummaryMainClass.ChangeData(YAxisCombinedOutputDNDLV.Items[i].Text);

            //Get Chart Summary Row
            InputFields1.PlotRowField = new string[XaxisRowDNDLV.Items.Count + ChartSeriesRowDNDLV.Items.Count];

            for (int i = 0; i < XaxisRowDNDLV.Items.Count; i++)
                InputFields1.PlotRowField[i] = SummaryMainClass.ChangeData(XaxisRowDNDLV.Items[i].Text);

            for (int i = 0; i < ChartSeriesRowDNDLV.Items.Count; i++)
                InputFields1.PlotRowField[i + XaxisRowDNDLV.Items.Count] = SummaryMainClass.ChangeData(ChartSeriesRowDNDLV.Items[i].Text);

            //Get Chart Summary Column
            InputFields1.PlotColumnField = new string[XaxisColumnDNDLV.Items.Count + ChartSeriesColumnDNDLV.Items.Count];

            for (int i = 0; i < XaxisColumnDNDLV.Items.Count; i++)
                InputFields1.PlotColumnField[i] = SummaryMainClass.ChangeData(XaxisColumnDNDLV.Items[i].Text);

            for (int i = 0; i < ChartSeriesColumnDNDLV.Items.Count; i++)
                InputFields1.PlotColumnField[i + XaxisColumnDNDLV.Items.Count] = SummaryMainClass.ChangeData(ChartSeriesColumnDNDLV.Items[i].Text);

            //Get Input Loop Iterator
            InputFields1.Iterator = new string[InputRowDNDLV.Items.Count + InputColumnDNDLV.Items.Count];

            for (int i = 0; i < InputRowDNDLV.Items.Count; i++)
                InputFields1.Iterator[i] = SummaryMainClass.ChangeData(InputRowDNDLV.Items[i].Text);

            for (int i = 0; i < InputColumnDNDLV.Items.Count; i++)
                InputFields1.Iterator[i + InputRowDNDLV.Items.Count] = SummaryMainClass.ChangeData(InputColumnDNDLV.Items[i].Text);

            //Get Datafield
            InputFields1.Datafield = new string[YAxisOutputDNDLV.Items.Count + YAxisCombinedOutputDNDLV.Items.Count];

            for (int i = 0; i < YAxisOutputDNDLV.Items.Count; i++)
                InputFields1.Datafield[i] = SummaryMainClass.ChangeData(YAxisOutputDNDLV.Items[i].Text);

            for (int i = 0; i < YAxisCombinedOutputDNDLV.Items.Count; i++)
                InputFields1.Datafield[i + YAxisOutputDNDLV.Items.Count] = SummaryMainClass.ChangeData(YAxisCombinedOutputDNDLV.Items[i].Text);

            //Get StatisticsDNDLV, StatisticsSelectedDNDLV
            InputFields1.Aggregate = new AggregateFunction[StatisticsSelectedDNDLV.Items.Count];

            for (int i = 0; i <= StatisticsSelectedDNDLV.Items.Count - 1; i++)
            {
                if (StatisticsSelectedDNDLV.Items[i].Text == AggregateFunction.Min.ToString()) InputFields1.Aggregate[i] = AggregateFunction.Min;
                if (StatisticsSelectedDNDLV.Items[i].Text == AggregateFunction.Max.ToString()) InputFields1.Aggregate[i] = AggregateFunction.Max;
                if (StatisticsSelectedDNDLV.Items[i].Text == AggregateFunction.Average.ToString()) InputFields1.Aggregate[i] = AggregateFunction.Average;
                if (StatisticsSelectedDNDLV.Items[i].Text == AggregateFunction.Stdev.ToString()) InputFields1.Aggregate[i] = AggregateFunction.Stdev;
                if (StatisticsSelectedDNDLV.Items[i].Text == AggregateFunction.NSigmaNeg.ToString()) InputFields1.Aggregate[i] = AggregateFunction.NSigmaNeg;
                if (StatisticsSelectedDNDLV.Items[i].Text == AggregateFunction.NSigmaPos.ToString()) InputFields1.Aggregate[i] = AggregateFunction.NSigmaPos;
                if (StatisticsSelectedDNDLV.Items[i].Text == AggregateFunction.SpecMin.ToString()) InputFields1.Aggregate[i] = AggregateFunction.SpecMin;
                if (StatisticsSelectedDNDLV.Items[i].Text == AggregateFunction.SpecMax.ToString()) InputFields1.Aggregate[i] = AggregateFunction.SpecMax;
                if (StatisticsSelectedDNDLV.Items[i].Text == AggregateFunction.SpecTypical.ToString()) InputFields1.Aggregate[i] = AggregateFunction.SpecTypical;
                if (StatisticsSelectedDNDLV.Items[i].Text == AggregateFunction.CPK.ToString()) InputFields1.Aggregate[i] = 
AggregateFunction.CPK;
                //if (StatisticsSelectedDNDLV.Items[i].Text == "") InputFields1.Aggregate[i] = AggregateFunction.Max;
            }
            List<AggregateFunction> lst_agr = InputFields1.Aggregate.ToList<AggregateFunction>();
            lst_agr.Remove(InputFields1.Aggregate.ToList<AggregateFunction>().Find(y => y.ToString().Equals("0")));
            InputFields1.Aggregate = lst_agr.ToArray();


            if (StatisticsSelectedDNDLV.Items.Count == 0) { InputFields1.Aggregate = new AggregateFunction[1]; InputFields1.Aggregate[0] = AggregateFunction.Max; }
            toolStripStatusLabel1.Text = "Ready";
            return true;
        }

        void AdminForminit()
        {
            releaseToQAToolStripMenuItem.Visible = false;
            releaseProductionToolStripMenuItem.Visible = false;
            notifyIcon1.Visible = false;
            toolStripProgressBar1.Visible = false;
            if (System.Security.Principal.WindowsIdentity.GetCurrent().Name == "SILABS_AD\\suandich")
            {
                releaseToQAToolStripMenuItem.Visible = true;
                releaseProductionToolStripMenuItem.Visible = true;
            }
        }
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Title title1 = new System.Windows.Forms.DataVisualization.Charting.Title();
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem("Min");
            System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem("Max");
            System.Windows.Forms.ListViewItem listViewItem3 = new System.Windows.Forms.ListViewItem("Average");
            System.Windows.Forms.ListViewItem listViewItem4 = new System.Windows.Forms.ListViewItem("Stdev");
            System.Windows.Forms.ListViewItem listViewItem5 = new System.Windows.Forms.ListViewItem("NSigmaPos");
            System.Windows.Forms.ListViewItem listViewItem6 = new System.Windows.Forms.ListViewItem("NSigmaNeg");
            System.Windows.Forms.ListViewItem listViewItem7 = new System.Windows.Forms.ListViewItem("CPK");
            System.Windows.Forms.ListViewItem listViewItem8 = new System.Windows.Forms.ListViewItem("SpecMin");
            System.Windows.Forms.ListViewItem listViewItem9 = new System.Windows.Forms.ListViewItem("SpecMax");
            System.Windows.Forms.ListViewItem listViewItem10 = new System.Windows.Forms.ListViewItem("SpecTypical");
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle13 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle15 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SummaryToolFrm));
            this.SumdataGridView = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menuToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.selectInputsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.restoreToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.createSpecToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exportToCSVToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.table1ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.table2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.chartSummaryToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.specificationNToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.highLevelSummaryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exportSummaryToPPTToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.chartSummaryPCToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.summaryPSToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AdminMode = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.disableSummaryToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.minimizeToTrayToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.disableChartDCToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.disableChartSummaryDUToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.disableInputSectionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.releaseToQAToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.releaseProductionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.updateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.revisionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.disableInputSectionToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.splitContainer3 = new System.Windows.Forms.SplitContainer();
            this.AllEntryDNDLV = new DragNDrop.DragNDropListView();
            this.columnHeader7 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader8 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.splitContainer4 = new System.Windows.Forms.SplitContainer();
            this.splitContainer5 = new System.Windows.Forms.SplitContainer();
            this.label5 = new System.Windows.Forms.Label();
            this.StatisticsSelectedDNDLV = new DragNDrop.DragNDropListView();
            this.columnHeader11 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader12 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.StatisticsDNDLV = new DragNDrop.DragNDropListView();
            this.columnHeader9 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader10 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.NSigmatextBox = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.menuStrip9 = new System.Windows.Forms.MenuStrip();
            this.statisticsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.splitContainer11 = new System.Windows.Forms.SplitContainer();
            this.splitContainer12 = new System.Windows.Forms.SplitContainer();
            this.XaxisRowDNDLV = new DragNDrop.DragNDropListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.menuStrip12 = new System.Windows.Forms.MenuStrip();
            this.xAxisRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.XaxisColumnDNDLV = new DragNDrop.DragNDropListView();
            this.columnHeader19 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader20 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.menuStrip14 = new System.Windows.Forms.MenuStrip();
            this.xAxisColumnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.splitContainer13 = new System.Windows.Forms.SplitContainer();
            this.ChartSeriesRowDNDLV = new DragNDrop.DragNDropListView();
            this.columnHeader17 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader18 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.menuStrip7 = new System.Windows.Forms.MenuStrip();
            this.chartSeriesRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ChartSeriesColumnDNDLV = new DragNDrop.DragNDropListView();
            this.columnHeader21 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader22 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.menuStrip13 = new System.Windows.Forms.MenuStrip();
            this.chartSeriesColumnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.splitContainer6 = new System.Windows.Forms.SplitContainer();
            this.splitContainer9 = new System.Windows.Forms.SplitContainer();
            this.InputRowDNDLV = new DragNDrop.DragNDropListView();
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader6 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.menuStrip6 = new System.Windows.Forms.MenuStrip();
            this.inputRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.InputColumnDNDLV = new DragNDrop.DragNDropListView();
            this.columnHeader13 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader14 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.menuStrip10 = new System.Windows.Forms.MenuStrip();
            this.inputColumnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.splitContainer10 = new System.Windows.Forms.SplitContainer();
            this.YAxisOutputDNDLV = new DragNDrop.DragNDropListView();
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.menuStrip5 = new System.Windows.Forms.MenuStrip();
            this.yAxisToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.YAxisCombinedOutputDNDLV = new DragNDrop.DragNDropListView();
            this.columnHeader15 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader16 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.menuStrip11 = new System.Windows.Forms.MenuStrip();
            this.multiOutputToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.splitContainer8 = new System.Windows.Forms.SplitContainer();
            this.splitContainer14 = new System.Windows.Forms.SplitContainer();
            this.HighSummaryDataGridView = new System.Windows.Forms.DataGridView();
            this.menuStrip2 = new System.Windows.Forms.MenuStrip();
            this.summaryTableToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.disableSummaryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.FilterdataGridView = new System.Windows.Forms.DataGridView();
            this.menuStrip3 = new System.Windows.Forms.MenuStrip();
            this.filterDataToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip4 = new System.Windows.Forms.MenuStrip();
            this.chartSummaryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.disableChartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.disableChartSummaryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.splitContainer7 = new System.Windows.Forms.SplitContainer();
            this.BpdataGridView = new System.Windows.Forms.DataGridView();
            this.menuStrip8 = new System.Windows.Forms.MenuStrip();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageSummaryTool = new System.Windows.Forms.TabPage();
            this.tabPageEdit = new System.Windows.Forms.TabPage();
            this.EditdataGridView = new System.Windows.Forms.DataGridView();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.SumdataGridView)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer3)).BeginInit();
            this.splitContainer3.Panel1.SuspendLayout();
            this.splitContainer3.Panel2.SuspendLayout();
            this.splitContainer3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer4)).BeginInit();
            this.splitContainer4.Panel1.SuspendLayout();
            this.splitContainer4.Panel2.SuspendLayout();
            this.splitContainer4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer5)).BeginInit();
            this.splitContainer5.Panel1.SuspendLayout();
            this.splitContainer5.Panel2.SuspendLayout();
            this.splitContainer5.SuspendLayout();
            this.menuStrip9.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer11)).BeginInit();
            this.splitContainer11.Panel1.SuspendLayout();
            this.splitContainer11.Panel2.SuspendLayout();
            this.splitContainer11.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer12)).BeginInit();
            this.splitContainer12.Panel1.SuspendLayout();
            this.splitContainer12.Panel2.SuspendLayout();
            this.splitContainer12.SuspendLayout();
            this.menuStrip12.SuspendLayout();
            this.menuStrip14.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer13)).BeginInit();
            this.splitContainer13.Panel1.SuspendLayout();
            this.splitContainer13.Panel2.SuspendLayout();
            this.splitContainer13.SuspendLayout();
            this.menuStrip7.SuspendLayout();
            this.menuStrip13.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer6)).BeginInit();
            this.splitContainer6.Panel1.SuspendLayout();
            this.splitContainer6.Panel2.SuspendLayout();
            this.splitContainer6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer9)).BeginInit();
            this.splitContainer9.Panel1.SuspendLayout();
            this.splitContainer9.Panel2.SuspendLayout();
            this.splitContainer9.SuspendLayout();
            this.menuStrip6.SuspendLayout();
            this.menuStrip10.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer10)).BeginInit();
            this.splitContainer10.Panel1.SuspendLayout();
            this.splitContainer10.Panel2.SuspendLayout();
            this.splitContainer10.SuspendLayout();
            this.menuStrip5.SuspendLayout();
            this.menuStrip11.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer8)).BeginInit();
            this.splitContainer8.Panel1.SuspendLayout();
            this.splitContainer8.Panel2.SuspendLayout();
            this.splitContainer8.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer14)).BeginInit();
            this.splitContainer14.Panel1.SuspendLayout();
            this.splitContainer14.Panel2.SuspendLayout();
            this.splitContainer14.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.HighSummaryDataGridView)).BeginInit();
            this.menuStrip2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.FilterdataGridView)).BeginInit();
            this.menuStrip3.SuspendLayout();
            this.menuStrip4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer7)).BeginInit();
            this.splitContainer7.Panel1.SuspendLayout();
            this.splitContainer7.Panel2.SuspendLayout();
            this.splitContainer7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.BpdataGridView)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPageSummaryTool.SuspendLayout();
            this.tabPageEdit.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.EditdataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // SumdataGridView
            // 
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.SumdataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.SumdataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.SumdataGridView.DefaultCellStyle = dataGridViewCellStyle2;
            this.SumdataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SumdataGridView.Location = new System.Drawing.Point(0, 0);
            this.SumdataGridView.Name = "SumdataGridView";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.SumdataGridView.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.SumdataGridView.Size = new System.Drawing.Size(805, 127);
            this.SumdataGridView.TabIndex = 2;
            this.SumdataGridView.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.SumdataGridView_CellContentClick);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuToolStripMenuItem,
            this.refreshToolStripMenuItem,
            this.releaseToQAToolStripMenuItem,
            this.releaseProductionToolStripMenuItem,
            this.updateToolStripMenuItem,
            this.disableInputSectionToolStripMenuItem1});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1276, 24);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // menuToolStripMenuItem
            // 
            this.menuToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.selectInputsToolStripMenuItem,
            this.restoreToolStripMenuItem,
            this.createSpecToolStripMenuItem,
            this.exportToCSVToolStripMenuItem,
            this.exportSummaryToPPTToolStripMenuItem,
            this.AdminMode,
            this.toolStripSeparator1,
            this.disableSummaryToolStripMenuItem1,
            this.minimizeToTrayToolStripMenuItem,
            this.disableChartDCToolStripMenuItem,
            this.disableChartSummaryDUToolStripMenuItem,
            this.disableInputSectionToolStripMenuItem,
            this.toolStripSeparator2,
            this.exitToolStripMenuItem});
            this.menuToolStripMenuItem.Name = "menuToolStripMenuItem";
            this.menuToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.menuToolStripMenuItem.Text = "&Menu";
            // 
            // selectInputsToolStripMenuItem
            // 
            this.selectInputsToolStripMenuItem.Name = "selectInputsToolStripMenuItem";
            this.selectInputsToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.selectInputsToolStripMenuItem.Text = "Select &Inputs                     I";
            this.selectInputsToolStripMenuItem.Click += new System.EventHandler(this.selectInputsToolStripMenuItem_Click);
            // 
            // restoreToolStripMenuItem
            // 
            this.restoreToolStripMenuItem.Name = "restoreToolStripMenuItem";
            this.restoreToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.restoreToolStripMenuItem.Text = "&Restore                              R";
            this.restoreToolStripMenuItem.Click += new System.EventHandler(this.restoreToolStripMenuItem_Click);
            // 
            // createSpecToolStripMenuItem
            // 
            this.createSpecToolStripMenuItem.Name = "createSpecToolStripMenuItem";
            this.createSpecToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.createSpecToolStripMenuItem.Text = "Recommanded &Spec       S";
            this.createSpecToolStripMenuItem.Click += new System.EventHandler(this.createSpecToolStripMenuItem_Click);
            // 
            // exportToCSVToolStripMenuItem
            // 
            this.exportToCSVToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.table1ToolStripMenuItem,
            this.highLevelSummaryToolStripMenuItem,
            this.table2ToolStripMenuItem,
            this.chartSummaryToolStripMenuItem1,
            this.specificationNToolStripMenuItem});
            this.exportToCSVToolStripMenuItem.Name = "exportToCSVToolStripMenuItem";
            this.exportToCSVToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.exportToCSVToolStripMenuItem.Text = "Export to &CSV";
            // 
            // table1ToolStripMenuItem
            // 
            this.table1ToolStripMenuItem.Name = "table1ToolStripMenuItem";
            this.table1ToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.table1ToolStripMenuItem.Text = "S&ummary                       U";
            this.table1ToolStripMenuItem.Click += new System.EventHandler(this.table1ToolStripMenuItem_Click);
            // 
            // table2ToolStripMenuItem
            // 
            this.table2ToolStripMenuItem.Name = "table2ToolStripMenuItem";
            this.table2ToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.table2ToolStripMenuItem.Text = "&Filtered Data                  F";
            this.table2ToolStripMenuItem.Click += new System.EventHandler(this.table2ToolStripMenuItem_Click);
            // 
            // chartSummaryToolStripMenuItem1
            // 
            this.chartSummaryToolStripMenuItem1.Name = "chartSummaryToolStripMenuItem1";
            this.chartSummaryToolStripMenuItem1.Size = new System.Drawing.Size(202, 22);
            this.chartSummaryToolStripMenuItem1.Text = "&Chart Summary            C";
            this.chartSummaryToolStripMenuItem1.Click += new System.EventHandler(this.chartSummaryToolStripMenuItem1_Click);
            // 
            // specificationNToolStripMenuItem
            // 
            this.specificationNToolStripMenuItem.Name = "specificationNToolStripMenuItem";
            this.specificationNToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.specificationNToolStripMenuItem.Text = "Specification                 N";
            // 
            // highLevelSummaryToolStripMenuItem
            // 
            this.highLevelSummaryToolStripMenuItem.Name = "highLevelSummaryToolStripMenuItem";
            this.highLevelSummaryToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.highLevelSummaryToolStripMenuItem.Text = "&HighLevel Summary    H";
            this.highLevelSummaryToolStripMenuItem.Click += new System.EventHandler(this.highLevelSummaryToolStripMenuItem_Click);
            // 
            // exportSummaryToPPTToolStripMenuItem
            // 
            this.exportSummaryToPPTToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.chartSummaryPCToolStripMenuItem,
            this.summaryPSToolStripMenuItem});
            this.exportSummaryToPPTToolStripMenuItem.Name = "exportSummaryToPPTToolStripMenuItem";
            this.exportSummaryToPPTToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.exportSummaryToPPTToolStripMenuItem.Text = "Export to PPT";
            // 
            // chartSummaryPCToolStripMenuItem
            // 
            this.chartSummaryPCToolStripMenuItem.Name = "chartSummaryPCToolStripMenuItem";
            this.chartSummaryPCToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.chartSummaryPCToolStripMenuItem.Text = "&Chart Summary    PC";
            this.chartSummaryPCToolStripMenuItem.Click += new System.EventHandler(this.chartSummaryPCToolStripMenuItem_Click);
            // 
            // summaryPSToolStripMenuItem
            // 
            this.summaryPSToolStripMenuItem.Name = "summaryPSToolStripMenuItem";
            this.summaryPSToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.summaryPSToolStripMenuItem.Text = "&Summary               PS";
            this.summaryPSToolStripMenuItem.Click += new System.EventHandler(this.summaryPSToolStripMenuItem_Click);
            // 
            // AdminMode
            // 
            this.AdminMode.Name = "AdminMode";
            this.AdminMode.Size = new System.Drawing.Size(220, 22);
            this.AdminMode.Text = "Admin Mode";
            this.AdminMode.Click += new System.EventHandler(this.AdminMode_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(217, 6);
            // 
            // disableSummaryToolStripMenuItem1
            // 
            this.disableSummaryToolStripMenuItem1.Name = "disableSummaryToolStripMenuItem1";
            this.disableSummaryToolStripMenuItem1.Size = new System.Drawing.Size(220, 22);
            this.disableSummaryToolStripMenuItem1.Text = "Disable Summary             DS";
            this.disableSummaryToolStripMenuItem1.Click += new System.EventHandler(this.disableSummaryToolStripMenuItem1_Click);
            // 
            // minimizeToTrayToolStripMenuItem
            // 
            this.minimizeToTrayToolStripMenuItem.Name = "minimizeToTrayToolStripMenuItem";
            this.minimizeToTrayToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.minimizeToTrayToolStripMenuItem.Text = "Minimize to Tray";
            this.minimizeToTrayToolStripMenuItem.Click += new System.EventHandler(this.minimizeToTrayToolStripMenuItem_Click);
            // 
            // disableChartDCToolStripMenuItem
            // 
            this.disableChartDCToolStripMenuItem.Name = "disableChartDCToolStripMenuItem";
            this.disableChartDCToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.disableChartDCToolStripMenuItem.Text = "Disable Chart                    DC";
            this.disableChartDCToolStripMenuItem.Click += new System.EventHandler(this.disableChartDCToolStripMenuItem_Click);
            // 
            // disableChartSummaryDUToolStripMenuItem
            // 
            this.disableChartSummaryDUToolStripMenuItem.Name = "disableChartSummaryDUToolStripMenuItem";
            this.disableChartSummaryDUToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.disableChartSummaryDUToolStripMenuItem.Text = "Disable Chart Summary  DU";
            this.disableChartSummaryDUToolStripMenuItem.Click += new System.EventHandler(this.disableChartSummaryDUToolStripMenuItem_Click);
            // 
            // disableInputSectionToolStripMenuItem
            // 
            this.disableInputSectionToolStripMenuItem.Name = "disableInputSectionToolStripMenuItem";
            this.disableInputSectionToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.disableInputSectionToolStripMenuItem.Text = "Disable Input Section      DI";
            this.disableInputSectionToolStripMenuItem.Click += new System.EventHandler(this.disableInputSectionToolStripMenuItem_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(217, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.exitToolStripMenuItem.Text = "E&xit                                      X";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // refreshToolStripMenuItem
            // 
            this.refreshToolStripMenuItem.Name = "refreshToolStripMenuItem";
            this.refreshToolStripMenuItem.Size = new System.Drawing.Size(73, 20);
            this.refreshToolStripMenuItem.Text = "&Refresh F5";
            this.refreshToolStripMenuItem.Click += new System.EventHandler(this.refreshToolStripMenuItem_Click);
            // 
            // releaseToQAToolStripMenuItem
            // 
            this.releaseToQAToolStripMenuItem.Name = "releaseToQAToolStripMenuItem";
            this.releaseToQAToolStripMenuItem.Size = new System.Drawing.Size(138, 20);
            this.releaseToQAToolStripMenuItem.Text = "Release to Engineering";
            this.releaseToQAToolStripMenuItem.Click += new System.EventHandler(this.releaseToQAToolStripMenuItem_Click);
            // 
            // releaseProductionToolStripMenuItem
            // 
            this.releaseProductionToolStripMenuItem.Name = "releaseProductionToolStripMenuItem";
            this.releaseProductionToolStripMenuItem.Size = new System.Drawing.Size(120, 20);
            this.releaseProductionToolStripMenuItem.Text = "Release Production";
            this.releaseProductionToolStripMenuItem.Click += new System.EventHandler(this.releaseProductionToolStripMenuItem_Click);
            // 
            // updateToolStripMenuItem
            // 
            this.updateToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.revisionToolStripMenuItem});
            this.updateToolStripMenuItem.Name = "updateToolStripMenuItem";
            this.updateToolStripMenuItem.Size = new System.Drawing.Size(57, 20);
            this.updateToolStripMenuItem.Text = "Update";
            this.updateToolStripMenuItem.Click += new System.EventHandler(this.updateToolStripMenuItem_Click);
            // 
            // revisionToolStripMenuItem
            // 
            this.revisionToolStripMenuItem.Name = "revisionToolStripMenuItem";
            this.revisionToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.revisionToolStripMenuItem.Text = "Revision";
            this.revisionToolStripMenuItem.Click += new System.EventHandler(this.revisionToolStripMenuItem_Click);
            // 
            // disableInputSectionToolStripMenuItem1
            // 
            this.disableInputSectionToolStripMenuItem1.Name = "disableInputSectionToolStripMenuItem1";
            this.disableInputSectionToolStripMenuItem1.Size = new System.Drawing.Size(130, 20);
            this.disableInputSectionToolStripMenuItem1.Text = "Disable Input Section";
            this.disableInputSectionToolStripMenuItem1.Click += new System.EventHandler(this.disableInputSectionToolStripMenuItem1_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 720);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1276, 22);
            this.statusStrip1.TabIndex = 4;
            this.statusStrip1.Text = "Ready";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(39, 17);
            this.toolStripStatusLabel1.Text = "Ready";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            // 
            // chart1
            // 
            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.chart1.Legends.Add(legend1);
            this.chart1.Location = new System.Drawing.Point(2, 27);
            this.chart1.Name = "chart1";
            series1.ChartArea = "ChartArea1";
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            this.chart1.Series.Add(series1);
            this.chart1.Size = new System.Drawing.Size(1046, 389);
            this.chart1.TabIndex = 5;
            this.chart1.Text = "chart1";
            title1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            title1.ForeColor = System.Drawing.Color.RoyalBlue;
            title1.Name = "Title1";
            title1.Text = "Box Plot";
            this.chart1.Titles.Add(title1);
            this.chart1.PostPaint += new System.EventHandler<System.Windows.Forms.DataVisualization.Charting.ChartPaintEventArgs>(this.chart1_PostPaint);
            this.chart1.Customize += new System.EventHandler(this.chart1_Customize);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(3, 3);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.splitContainer3);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(1262, 664);
            this.splitContainer1.SplitterDistance = 185;
            this.splitContainer1.TabIndex = 6;
            // 
            // splitContainer3
            // 
            this.splitContainer3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer3.Location = new System.Drawing.Point(0, 0);
            this.splitContainer3.Name = "splitContainer3";
            // 
            // splitContainer3.Panel1
            // 
            this.splitContainer3.Panel1.Controls.Add(this.AllEntryDNDLV);
            // 
            // splitContainer3.Panel2
            // 
            this.splitContainer3.Panel2.Controls.Add(this.splitContainer4);
            this.splitContainer3.Size = new System.Drawing.Size(1262, 185);
            this.splitContainer3.SplitterDistance = 385;
            this.splitContainer3.TabIndex = 2;
            // 
            // AllEntryDNDLV
            // 
            this.AllEntryDNDLV.AllowDrop = true;
            this.AllEntryDNDLV.AllowReorder = true;
            this.AllEntryDNDLV.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader7,
            this.columnHeader8});
            this.AllEntryDNDLV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.AllEntryDNDLV.FullRowSelect = true;
            this.AllEntryDNDLV.LineColor = System.Drawing.Color.Olive;
            this.AllEntryDNDLV.Location = new System.Drawing.Point(0, 0);
            this.AllEntryDNDLV.Name = "AllEntryDNDLV";
            this.AllEntryDNDLV.Size = new System.Drawing.Size(385, 185);
            this.AllEntryDNDLV.TabIndex = 1;
            this.AllEntryDNDLV.UseCompatibleStateImageBehavior = false;
            this.AllEntryDNDLV.View = System.Windows.Forms.View.List;
            this.AllEntryDNDLV.DragDrop += new System.Windows.Forms.DragEventHandler(this.AllEntryDNDLV_DragDrop);
            this.AllEntryDNDLV.MouseClick += new System.Windows.Forms.MouseEventHandler(this.AllEntryDNDLV_MouseClick);
            // 
            // columnHeader7
            // 
            this.columnHeader7.Width = 190;
            // 
            // columnHeader8
            // 
            this.columnHeader8.Width = 244;
            // 
            // splitContainer4
            // 
            this.splitContainer4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer4.Location = new System.Drawing.Point(0, 0);
            this.splitContainer4.Name = "splitContainer4";
            this.splitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer4.Panel1
            // 
            this.splitContainer4.Panel1.Controls.Add(this.splitContainer5);
            // 
            // splitContainer4.Panel2
            // 
            this.splitContainer4.Panel2.Controls.Add(this.splitContainer6);
            this.splitContainer4.Size = new System.Drawing.Size(873, 185);
            this.splitContainer4.SplitterDistance = 109;
            this.splitContainer4.TabIndex = 0;
            // 
            // splitContainer5
            // 
            this.splitContainer5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer5.Location = new System.Drawing.Point(0, 0);
            this.splitContainer5.Name = "splitContainer5";
            // 
            // splitContainer5.Panel1
            // 
            this.splitContainer5.Panel1.Controls.Add(this.label5);
            this.splitContainer5.Panel1.Controls.Add(this.StatisticsSelectedDNDLV);
            this.splitContainer5.Panel1.Controls.Add(this.StatisticsDNDLV);
            this.splitContainer5.Panel1.Controls.Add(this.NSigmatextBox);
            this.splitContainer5.Panel1.Controls.Add(this.label4);
            this.splitContainer5.Panel1.Controls.Add(this.menuStrip9);
            // 
            // splitContainer5.Panel2
            // 
            this.splitContainer5.Panel2.Controls.Add(this.splitContainer11);
            this.splitContainer5.Size = new System.Drawing.Size(873, 109);
            this.splitContainer5.SplitterDistance = 373;
            this.splitContainer5.TabIndex = 0;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(102, 61);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(16, 13);
            this.label5.TabIndex = 5;
            this.label5.Text = "->";
            // 
            // StatisticsSelectedDNDLV
            // 
            this.StatisticsSelectedDNDLV.AllowDrop = true;
            this.StatisticsSelectedDNDLV.AllowReorder = true;
            this.StatisticsSelectedDNDLV.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader11,
            this.columnHeader12});
            this.StatisticsSelectedDNDLV.FullRowSelect = true;
            this.StatisticsSelectedDNDLV.LineColor = System.Drawing.Color.Olive;
            this.StatisticsSelectedDNDLV.Location = new System.Drawing.Point(124, 24);
            this.StatisticsSelectedDNDLV.Name = "StatisticsSelectedDNDLV";
            this.StatisticsSelectedDNDLV.Size = new System.Drawing.Size(93, 112);
            this.StatisticsSelectedDNDLV.TabIndex = 1;
            this.StatisticsSelectedDNDLV.UseCompatibleStateImageBehavior = false;
            this.StatisticsSelectedDNDLV.View = System.Windows.Forms.View.List;
            this.StatisticsSelectedDNDLV.DragDrop += new System.Windows.Forms.DragEventHandler(this.StatisticsSelectedDNDLV_DragDrop);
            // 
            // columnHeader11
            // 
            this.columnHeader11.Width = 190;
            // 
            // columnHeader12
            // 
            this.columnHeader12.Width = 244;
            // 
            // StatisticsDNDLV
            // 
            this.StatisticsDNDLV.AllowDrop = true;
            this.StatisticsDNDLV.AllowReorder = true;
            this.StatisticsDNDLV.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader9,
            this.columnHeader10});
            this.StatisticsDNDLV.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.StatisticsDNDLV.FullRowSelect = true;
            this.StatisticsDNDLV.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1,
            listViewItem2,
            listViewItem3,
            listViewItem4,
            listViewItem5,
            listViewItem6,
            listViewItem7,
            listViewItem8,
            listViewItem9,
            listViewItem10});
            this.StatisticsDNDLV.LineColor = System.Drawing.Color.Olive;
            this.StatisticsDNDLV.Location = new System.Drawing.Point(3, 24);
            this.StatisticsDNDLV.Name = "StatisticsDNDLV";
            this.StatisticsDNDLV.Size = new System.Drawing.Size(93, 112);
            this.StatisticsDNDLV.TabIndex = 1;
            this.StatisticsDNDLV.UseCompatibleStateImageBehavior = false;
            this.StatisticsDNDLV.View = System.Windows.Forms.View.List;
            this.StatisticsDNDLV.DragDrop += new System.Windows.Forms.DragEventHandler(this.StatisticsDNDLV_DragDrop);
            // 
            // columnHeader9
            // 
            this.columnHeader9.Width = 190;
            // 
            // columnHeader10
            // 
            this.columnHeader10.Width = 244;
            // 
            // NSigmatextBox
            // 
            this.NSigmatextBox.Location = new System.Drawing.Point(231, 103);
            this.NSigmatextBox.Name = "NSigmatextBox";
            this.NSigmatextBox.Size = new System.Drawing.Size(63, 20);
            this.NSigmatextBox.TabIndex = 4;
            this.NSigmatextBox.TextChanged += new System.EventHandler(this.NSigmatextBox_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(228, 87);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(66, 13);
            this.label4.TabIndex = 2;
            this.label4.Text = "N in NSigma";
            // 
            // menuStrip9
            // 
            this.menuStrip9.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statisticsToolStripMenuItem});
            this.menuStrip9.Location = new System.Drawing.Point(0, 0);
            this.menuStrip9.Name = "menuStrip9";
            this.menuStrip9.Size = new System.Drawing.Size(373, 24);
            this.menuStrip9.TabIndex = 6;
            this.menuStrip9.Text = "menuStrip9";
            // 
            // statisticsToolStripMenuItem
            // 
            this.statisticsToolStripMenuItem.Enabled = false;
            this.statisticsToolStripMenuItem.Name = "statisticsToolStripMenuItem";
            this.statisticsToolStripMenuItem.Size = new System.Drawing.Size(65, 20);
            this.statisticsToolStripMenuItem.Text = "Statistics";
            // 
            // splitContainer11
            // 
            this.splitContainer11.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer11.Location = new System.Drawing.Point(0, 0);
            this.splitContainer11.Name = "splitContainer11";
            // 
            // splitContainer11.Panel1
            // 
            this.splitContainer11.Panel1.Controls.Add(this.splitContainer12);
            // 
            // splitContainer11.Panel2
            // 
            this.splitContainer11.Panel2.Controls.Add(this.splitContainer13);
            this.splitContainer11.Size = new System.Drawing.Size(496, 109);
            this.splitContainer11.SplitterDistance = 191;
            this.splitContainer11.TabIndex = 2;
            // 
            // splitContainer12
            // 
            this.splitContainer12.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer12.Location = new System.Drawing.Point(0, 0);
            this.splitContainer12.Name = "splitContainer12";
            // 
            // splitContainer12.Panel1
            // 
            this.splitContainer12.Panel1.Controls.Add(this.XaxisRowDNDLV);
            this.splitContainer12.Panel1.Controls.Add(this.menuStrip12);
            // 
            // splitContainer12.Panel2
            // 
            this.splitContainer12.Panel2.Controls.Add(this.XaxisColumnDNDLV);
            this.splitContainer12.Panel2.Controls.Add(this.menuStrip14);
            this.splitContainer12.Size = new System.Drawing.Size(191, 109);
            this.splitContainer12.SplitterDistance = 88;
            this.splitContainer12.TabIndex = 2;
            // 
            // XaxisRowDNDLV
            // 
            this.XaxisRowDNDLV.AllowDrop = true;
            this.XaxisRowDNDLV.AllowReorder = true;
            this.XaxisRowDNDLV.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2});
            this.XaxisRowDNDLV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.XaxisRowDNDLV.FullRowSelect = true;
            this.XaxisRowDNDLV.LineColor = System.Drawing.Color.Olive;
            this.XaxisRowDNDLV.Location = new System.Drawing.Point(0, 24);
            this.XaxisRowDNDLV.Name = "XaxisRowDNDLV";
            this.XaxisRowDNDLV.Size = new System.Drawing.Size(88, 85);
            this.XaxisRowDNDLV.TabIndex = 0;
            this.XaxisRowDNDLV.UseCompatibleStateImageBehavior = false;
            this.XaxisRowDNDLV.View = System.Windows.Forms.View.List;
            this.XaxisRowDNDLV.DragDrop += new System.Windows.Forms.DragEventHandler(this.XaxisRowDNDLV_DragDrop);
            this.XaxisRowDNDLV.MouseClick += new System.Windows.Forms.MouseEventHandler(this.XaxisRowDNDLV_MouseClick);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Width = 190;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Width = 244;
            // 
            // menuStrip12
            // 
            this.menuStrip12.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.xAxisRowToolStripMenuItem});
            this.menuStrip12.Location = new System.Drawing.Point(0, 0);
            this.menuStrip12.Name = "menuStrip12";
            this.menuStrip12.Size = new System.Drawing.Size(88, 24);
            this.menuStrip12.TabIndex = 1;
            this.menuStrip12.Text = "menuStrip12";
            // 
            // xAxisRowToolStripMenuItem
            // 
            this.xAxisRowToolStripMenuItem.Enabled = false;
            this.xAxisRowToolStripMenuItem.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.xAxisRowToolStripMenuItem.Name = "xAxisRowToolStripMenuItem";
            this.xAxisRowToolStripMenuItem.Size = new System.Drawing.Size(74, 20);
            this.xAxisRowToolStripMenuItem.Text = "X Axis Row";
            // 
            // XaxisColumnDNDLV
            // 
            this.XaxisColumnDNDLV.AllowDrop = true;
            this.XaxisColumnDNDLV.AllowReorder = true;
            this.XaxisColumnDNDLV.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader19,
            this.columnHeader20});
            this.XaxisColumnDNDLV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.XaxisColumnDNDLV.FullRowSelect = true;
            this.XaxisColumnDNDLV.LineColor = System.Drawing.Color.Olive;
            this.XaxisColumnDNDLV.Location = new System.Drawing.Point(0, 24);
            this.XaxisColumnDNDLV.Name = "XaxisColumnDNDLV";
            this.XaxisColumnDNDLV.Size = new System.Drawing.Size(99, 85);
            this.XaxisColumnDNDLV.TabIndex = 0;
            this.XaxisColumnDNDLV.UseCompatibleStateImageBehavior = false;
            this.XaxisColumnDNDLV.View = System.Windows.Forms.View.List;
            this.XaxisColumnDNDLV.DragDrop += new System.Windows.Forms.DragEventHandler(this.XaxisColumnDNDLV_DragDrop);
            this.XaxisColumnDNDLV.MouseClick += new System.Windows.Forms.MouseEventHandler(this.XaxisColumnDNDLV_MouseClick);
            // 
            // columnHeader19
            // 
            this.columnHeader19.Width = 190;
            // 
            // columnHeader20
            // 
            this.columnHeader20.Width = 244;
            // 
            // menuStrip14
            // 
            this.menuStrip14.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.xAxisColumnToolStripMenuItem});
            this.menuStrip14.Location = new System.Drawing.Point(0, 0);
            this.menuStrip14.Name = "menuStrip14";
            this.menuStrip14.Size = new System.Drawing.Size(99, 24);
            this.menuStrip14.TabIndex = 1;
            this.menuStrip14.Text = "menuStrip14";
            // 
            // xAxisColumnToolStripMenuItem
            // 
            this.xAxisColumnToolStripMenuItem.Enabled = false;
            this.xAxisColumnToolStripMenuItem.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.xAxisColumnToolStripMenuItem.Name = "xAxisColumnToolStripMenuItem";
            this.xAxisColumnToolStripMenuItem.Size = new System.Drawing.Size(88, 20);
            this.xAxisColumnToolStripMenuItem.Text = "X AxisColumn";
            // 
            // splitContainer13
            // 
            this.splitContainer13.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer13.Location = new System.Drawing.Point(0, 0);
            this.splitContainer13.Name = "splitContainer13";
            // 
            // splitContainer13.Panel1
            // 
            this.splitContainer13.Panel1.Controls.Add(this.ChartSeriesRowDNDLV);
            this.splitContainer13.Panel1.Controls.Add(this.menuStrip7);
            // 
            // splitContainer13.Panel2
            // 
            this.splitContainer13.Panel2.Controls.Add(this.ChartSeriesColumnDNDLV);
            this.splitContainer13.Panel2.Controls.Add(this.menuStrip13);
            this.splitContainer13.Size = new System.Drawing.Size(301, 109);
            this.splitContainer13.SplitterDistance = 122;
            this.splitContainer13.TabIndex = 2;
            // 
            // ChartSeriesRowDNDLV
            // 
            this.ChartSeriesRowDNDLV.AllowDrop = true;
            this.ChartSeriesRowDNDLV.AllowReorder = true;
            this.ChartSeriesRowDNDLV.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader17,
            this.columnHeader18});
            this.ChartSeriesRowDNDLV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ChartSeriesRowDNDLV.FullRowSelect = true;
            this.ChartSeriesRowDNDLV.LineColor = System.Drawing.Color.Olive;
            this.ChartSeriesRowDNDLV.Location = new System.Drawing.Point(0, 24);
            this.ChartSeriesRowDNDLV.Name = "ChartSeriesRowDNDLV";
            this.ChartSeriesRowDNDLV.Size = new System.Drawing.Size(122, 85);
            this.ChartSeriesRowDNDLV.TabIndex = 0;
            this.ChartSeriesRowDNDLV.UseCompatibleStateImageBehavior = false;
            this.ChartSeriesRowDNDLV.View = System.Windows.Forms.View.List;
            this.ChartSeriesRowDNDLV.DragDrop += new System.Windows.Forms.DragEventHandler(this.ChartSeriesRowDNDLV_DragDrop);
            this.ChartSeriesRowDNDLV.MouseClick += new System.Windows.Forms.MouseEventHandler(this.ChartSeriesRowDNDLV_MouseClick);
            // 
            // columnHeader17
            // 
            this.columnHeader17.Width = 190;
            // 
            // columnHeader18
            // 
            this.columnHeader18.Width = 244;
            // 
            // menuStrip7
            // 
            this.menuStrip7.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.chartSeriesRowToolStripMenuItem});
            this.menuStrip7.Location = new System.Drawing.Point(0, 0);
            this.menuStrip7.Name = "menuStrip7";
            this.menuStrip7.Size = new System.Drawing.Size(122, 24);
            this.menuStrip7.TabIndex = 1;
            this.menuStrip7.Text = "menuStrip7";
            // 
            // chartSeriesRowToolStripMenuItem
            // 
            this.chartSeriesRowToolStripMenuItem.Enabled = false;
            this.chartSeriesRowToolStripMenuItem.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chartSeriesRowToolStripMenuItem.Name = "chartSeriesRowToolStripMenuItem";
            this.chartSeriesRowToolStripMenuItem.Size = new System.Drawing.Size(106, 20);
            this.chartSeriesRowToolStripMenuItem.Text = "Chart Series Row";
            // 
            // ChartSeriesColumnDNDLV
            // 
            this.ChartSeriesColumnDNDLV.AllowDrop = true;
            this.ChartSeriesColumnDNDLV.AllowReorder = true;
            this.ChartSeriesColumnDNDLV.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader21,
            this.columnHeader22});
            this.ChartSeriesColumnDNDLV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ChartSeriesColumnDNDLV.FullRowSelect = true;
            this.ChartSeriesColumnDNDLV.LineColor = System.Drawing.Color.Olive;
            this.ChartSeriesColumnDNDLV.Location = new System.Drawing.Point(0, 24);
            this.ChartSeriesColumnDNDLV.Name = "ChartSeriesColumnDNDLV";
            this.ChartSeriesColumnDNDLV.Size = new System.Drawing.Size(175, 85);
            this.ChartSeriesColumnDNDLV.TabIndex = 0;
            this.ChartSeriesColumnDNDLV.UseCompatibleStateImageBehavior = false;
            this.ChartSeriesColumnDNDLV.View = System.Windows.Forms.View.List;
            this.ChartSeriesColumnDNDLV.DragDrop += new System.Windows.Forms.DragEventHandler(this.ChartSeriesColumnDNDLV_DragDrop);
            this.ChartSeriesColumnDNDLV.MouseClick += new System.Windows.Forms.MouseEventHandler(this.ChartSeriesColumnDNDLV_MouseClick);
            // 
            // columnHeader21
            // 
            this.columnHeader21.Width = 190;
            // 
            // columnHeader22
            // 
            this.columnHeader22.Width = 244;
            // 
            // menuStrip13
            // 
            this.menuStrip13.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.chartSeriesColumnToolStripMenuItem});
            this.menuStrip13.Location = new System.Drawing.Point(0, 0);
            this.menuStrip13.Name = "menuStrip13";
            this.menuStrip13.Size = new System.Drawing.Size(175, 24);
            this.menuStrip13.TabIndex = 1;
            this.menuStrip13.Text = "menuStrip13";
            // 
            // chartSeriesColumnToolStripMenuItem
            // 
            this.chartSeriesColumnToolStripMenuItem.Enabled = false;
            this.chartSeriesColumnToolStripMenuItem.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chartSeriesColumnToolStripMenuItem.Name = "chartSeriesColumnToolStripMenuItem";
            this.chartSeriesColumnToolStripMenuItem.Size = new System.Drawing.Size(123, 20);
            this.chartSeriesColumnToolStripMenuItem.Text = "Chart Series Column";
            // 
            // splitContainer6
            // 
            this.splitContainer6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer6.Location = new System.Drawing.Point(0, 0);
            this.splitContainer6.Name = "splitContainer6";
            // 
            // splitContainer6.Panel1
            // 
            this.splitContainer6.Panel1.Controls.Add(this.splitContainer9);
            // 
            // splitContainer6.Panel2
            // 
            this.splitContainer6.Panel2.Controls.Add(this.splitContainer10);
            this.splitContainer6.Size = new System.Drawing.Size(873, 72);
            this.splitContainer6.SplitterDistance = 373;
            this.splitContainer6.TabIndex = 0;
            // 
            // splitContainer9
            // 
            this.splitContainer9.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer9.Location = new System.Drawing.Point(0, 0);
            this.splitContainer9.Name = "splitContainer9";
            // 
            // splitContainer9.Panel1
            // 
            this.splitContainer9.Panel1.Controls.Add(this.InputRowDNDLV);
            this.splitContainer9.Panel1.Controls.Add(this.menuStrip6);
            // 
            // splitContainer9.Panel2
            // 
            this.splitContainer9.Panel2.Controls.Add(this.InputColumnDNDLV);
            this.splitContainer9.Panel2.Controls.Add(this.menuStrip10);
            this.splitContainer9.Size = new System.Drawing.Size(373, 72);
            this.splitContainer9.SplitterDistance = 177;
            this.splitContainer9.TabIndex = 3;
            // 
            // InputRowDNDLV
            // 
            this.InputRowDNDLV.AllowDrop = true;
            this.InputRowDNDLV.AllowReorder = true;
            this.InputRowDNDLV.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader5,
            this.columnHeader6});
            this.InputRowDNDLV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.InputRowDNDLV.FullRowSelect = true;
            this.InputRowDNDLV.LineColor = System.Drawing.Color.Olive;
            this.InputRowDNDLV.Location = new System.Drawing.Point(0, 24);
            this.InputRowDNDLV.Name = "InputRowDNDLV";
            this.InputRowDNDLV.Size = new System.Drawing.Size(177, 48);
            this.InputRowDNDLV.TabIndex = 1;
            this.InputRowDNDLV.UseCompatibleStateImageBehavior = false;
            this.InputRowDNDLV.View = System.Windows.Forms.View.List;
            this.InputRowDNDLV.DragDrop += new System.Windows.Forms.DragEventHandler(this.InputRowDNDLV_DragDrop);
            this.InputRowDNDLV.MouseClick += new System.Windows.Forms.MouseEventHandler(this.InputRowDNDLV_MouseClick);
            // 
            // columnHeader5
            // 
            this.columnHeader5.Width = 190;
            // 
            // columnHeader6
            // 
            this.columnHeader6.Width = 244;
            // 
            // menuStrip6
            // 
            this.menuStrip6.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.inputRowToolStripMenuItem});
            this.menuStrip6.Location = new System.Drawing.Point(0, 0);
            this.menuStrip6.Name = "menuStrip6";
            this.menuStrip6.Size = new System.Drawing.Size(177, 24);
            this.menuStrip6.TabIndex = 2;
            this.menuStrip6.Text = "menuStrip6";
            // 
            // inputRowToolStripMenuItem
            // 
            this.inputRowToolStripMenuItem.AutoToolTip = true;
            this.inputRowToolStripMenuItem.Enabled = false;
            this.inputRowToolStripMenuItem.Name = "inputRowToolStripMenuItem";
            this.inputRowToolStripMenuItem.Size = new System.Drawing.Size(73, 20);
            this.inputRowToolStripMenuItem.Text = "Input Row";
            this.inputRowToolStripMenuItem.ToolTipText = "Select List of Inputs to iterate through each entry / Row in the summary table";
            // 
            // InputColumnDNDLV
            // 
            this.InputColumnDNDLV.AllowDrop = true;
            this.InputColumnDNDLV.AllowReorder = true;
            this.InputColumnDNDLV.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader13,
            this.columnHeader14});
            this.InputColumnDNDLV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.InputColumnDNDLV.FullRowSelect = true;
            this.InputColumnDNDLV.LineColor = System.Drawing.Color.Olive;
            this.InputColumnDNDLV.Location = new System.Drawing.Point(0, 24);
            this.InputColumnDNDLV.Name = "InputColumnDNDLV";
            this.InputColumnDNDLV.Size = new System.Drawing.Size(192, 48);
            this.InputColumnDNDLV.TabIndex = 1;
            this.InputColumnDNDLV.UseCompatibleStateImageBehavior = false;
            this.InputColumnDNDLV.View = System.Windows.Forms.View.List;
            this.InputColumnDNDLV.DragDrop += new System.Windows.Forms.DragEventHandler(this.InputColumnDNDLV_DragDrop);
            this.InputColumnDNDLV.MouseClick += new System.Windows.Forms.MouseEventHandler(this.InputColumnDNDLV_MouseClick);
            // 
            // columnHeader13
            // 
            this.columnHeader13.Width = 190;
            // 
            // columnHeader14
            // 
            this.columnHeader14.Width = 244;
            // 
            // menuStrip10
            // 
            this.menuStrip10.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.inputColumnToolStripMenuItem});
            this.menuStrip10.Location = new System.Drawing.Point(0, 0);
            this.menuStrip10.Name = "menuStrip10";
            this.menuStrip10.Size = new System.Drawing.Size(192, 24);
            this.menuStrip10.TabIndex = 2;
            this.menuStrip10.Text = "menuStrip10";
            // 
            // inputColumnToolStripMenuItem
            // 
            this.inputColumnToolStripMenuItem.Enabled = false;
            this.inputColumnToolStripMenuItem.Name = "inputColumnToolStripMenuItem";
            this.inputColumnToolStripMenuItem.Size = new System.Drawing.Size(93, 20);
            this.inputColumnToolStripMenuItem.Text = "Input Column";
            this.inputColumnToolStripMenuItem.ToolTipText = "Select List of Inputs to iterate through each entry / Column in the summary table" +
                "";
            // 
            // splitContainer10
            // 
            this.splitContainer10.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer10.Location = new System.Drawing.Point(0, 0);
            this.splitContainer10.Name = "splitContainer10";
            // 
            // splitContainer10.Panel1
            // 
            this.splitContainer10.Panel1.Controls.Add(this.YAxisOutputDNDLV);
            this.splitContainer10.Panel1.Controls.Add(this.menuStrip5);
            // 
            // splitContainer10.Panel2
            // 
            this.splitContainer10.Panel2.Controls.Add(this.YAxisCombinedOutputDNDLV);
            this.splitContainer10.Panel2.Controls.Add(this.menuStrip11);
            this.splitContainer10.Size = new System.Drawing.Size(496, 72);
            this.splitContainer10.SplitterDistance = 234;
            this.splitContainer10.TabIndex = 2;
            // 
            // YAxisOutputDNDLV
            // 
            this.YAxisOutputDNDLV.AllowDrop = true;
            this.YAxisOutputDNDLV.AllowReorder = true;
            this.YAxisOutputDNDLV.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader3,
            this.columnHeader4});
            this.YAxisOutputDNDLV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.YAxisOutputDNDLV.FullRowSelect = true;
            this.YAxisOutputDNDLV.LineColor = System.Drawing.Color.Olive;
            this.YAxisOutputDNDLV.Location = new System.Drawing.Point(0, 24);
            this.YAxisOutputDNDLV.Name = "YAxisOutputDNDLV";
            this.YAxisOutputDNDLV.Size = new System.Drawing.Size(234, 48);
            this.YAxisOutputDNDLV.TabIndex = 1;
            this.YAxisOutputDNDLV.UseCompatibleStateImageBehavior = false;
            this.YAxisOutputDNDLV.View = System.Windows.Forms.View.List;
            this.YAxisOutputDNDLV.DragDrop += new System.Windows.Forms.DragEventHandler(this.YAxisOutputDNDLV_DragDrop_1);
            // 
            // columnHeader3
            // 
            this.columnHeader3.Width = 190;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Width = 244;
            // 
            // menuStrip5
            // 
            this.menuStrip5.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.yAxisToolStripMenuItem});
            this.menuStrip5.Location = new System.Drawing.Point(0, 0);
            this.menuStrip5.Name = "menuStrip5";
            this.menuStrip5.Size = new System.Drawing.Size(234, 24);
            this.menuStrip5.TabIndex = 2;
            this.menuStrip5.Text = "menuStrip5";
            // 
            // yAxisToolStripMenuItem
            // 
            this.yAxisToolStripMenuItem.Enabled = false;
            this.yAxisToolStripMenuItem.Name = "yAxisToolStripMenuItem";
            this.yAxisToolStripMenuItem.Size = new System.Drawing.Size(91, 20);
            this.yAxisToolStripMenuItem.Text = "Y Axis Output";
            // 
            // YAxisCombinedOutputDNDLV
            // 
            this.YAxisCombinedOutputDNDLV.AllowDrop = true;
            this.YAxisCombinedOutputDNDLV.AllowReorder = true;
            this.YAxisCombinedOutputDNDLV.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader15,
            this.columnHeader16});
            this.YAxisCombinedOutputDNDLV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.YAxisCombinedOutputDNDLV.FullRowSelect = true;
            this.YAxisCombinedOutputDNDLV.LineColor = System.Drawing.Color.Olive;
            this.YAxisCombinedOutputDNDLV.Location = new System.Drawing.Point(0, 24);
            this.YAxisCombinedOutputDNDLV.Name = "YAxisCombinedOutputDNDLV";
            this.YAxisCombinedOutputDNDLV.Size = new System.Drawing.Size(258, 48);
            this.YAxisCombinedOutputDNDLV.TabIndex = 1;
            this.YAxisCombinedOutputDNDLV.UseCompatibleStateImageBehavior = false;
            this.YAxisCombinedOutputDNDLV.View = System.Windows.Forms.View.List;
            this.YAxisCombinedOutputDNDLV.DragDrop += new System.Windows.Forms.DragEventHandler(this.YAxisCombinedOutputDNDLV_DragDrop);
            // 
            // columnHeader15
            // 
            this.columnHeader15.Width = 190;
            // 
            // columnHeader16
            // 
            this.columnHeader16.Width = 244;
            // 
            // menuStrip11
            // 
            this.menuStrip11.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.multiOutputToolStripMenuItem});
            this.menuStrip11.Location = new System.Drawing.Point(0, 0);
            this.menuStrip11.Name = "menuStrip11";
            this.menuStrip11.Size = new System.Drawing.Size(258, 24);
            this.menuStrip11.TabIndex = 2;
            this.menuStrip11.Text = "menuStrip11";
            // 
            // multiOutputToolStripMenuItem
            // 
            this.multiOutputToolStripMenuItem.Enabled = false;
            this.multiOutputToolStripMenuItem.Name = "multiOutputToolStripMenuItem";
            this.multiOutputToolStripMenuItem.Size = new System.Drawing.Size(150, 20);
            this.multiOutputToolStripMenuItem.Text = "Y Axis Combined Output";
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            this.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.splitContainer8);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.menuStrip4);
            this.splitContainer2.Panel2.Controls.Add(this.splitContainer7);
            this.splitContainer2.Size = new System.Drawing.Size(1262, 475);
            this.splitContainer2.SplitterDistance = 196;
            this.splitContainer2.TabIndex = 0;
            // 
            // splitContainer8
            // 
            this.splitContainer8.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer8.Location = new System.Drawing.Point(0, 0);
            this.splitContainer8.Name = "splitContainer8";
            // 
            // splitContainer8.Panel1
            // 
            this.splitContainer8.Panel1.Controls.Add(this.splitContainer14);
            this.splitContainer8.Panel1.Controls.Add(this.menuStrip2);
            // 
            // splitContainer8.Panel2
            // 
            this.splitContainer8.Panel2.Controls.Add(this.FilterdataGridView);
            this.splitContainer8.Panel2.Controls.Add(this.menuStrip3);
            this.splitContainer8.Size = new System.Drawing.Size(1262, 196);
            this.splitContainer8.SplitterDistance = 805;
            this.splitContainer8.TabIndex = 3;
            // 
            // splitContainer14
            // 
            this.splitContainer14.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer14.Location = new System.Drawing.Point(0, 24);
            this.splitContainer14.Name = "splitContainer14";
            this.splitContainer14.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer14.Panel1
            // 
            this.splitContainer14.Panel1.Controls.Add(this.SumdataGridView);
            // 
            // splitContainer14.Panel2
            // 
            this.splitContainer14.Panel2.Controls.Add(this.HighSummaryDataGridView);
            this.splitContainer14.Size = new System.Drawing.Size(805, 172);
            this.splitContainer14.SplitterDistance = 127;
            this.splitContainer14.TabIndex = 4;
            // 
            // HighSummaryDataGridView
            // 
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.HighSummaryDataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.HighSummaryDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.HighSummaryDataGridView.DefaultCellStyle = dataGridViewCellStyle5;
            this.HighSummaryDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.HighSummaryDataGridView.Location = new System.Drawing.Point(0, 0);
            this.HighSummaryDataGridView.Name = "HighSummaryDataGridView";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.HighSummaryDataGridView.RowHeadersDefaultCellStyle = dataGridViewCellStyle6;
            this.HighSummaryDataGridView.Size = new System.Drawing.Size(805, 41);
            this.HighSummaryDataGridView.TabIndex = 3;
            // 
            // menuStrip2
            // 
            this.menuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.summaryTableToolStripMenuItem,
            this.disableSummaryToolStripMenuItem});
            this.menuStrip2.Location = new System.Drawing.Point(0, 0);
            this.menuStrip2.Name = "menuStrip2";
            this.menuStrip2.Size = new System.Drawing.Size(805, 24);
            this.menuStrip2.TabIndex = 3;
            this.menuStrip2.Text = "menuStrip2";
            // 
            // summaryTableToolStripMenuItem
            // 
            this.summaryTableToolStripMenuItem.Enabled = false;
            this.summaryTableToolStripMenuItem.Name = "summaryTableToolStripMenuItem";
            this.summaryTableToolStripMenuItem.Size = new System.Drawing.Size(102, 20);
            this.summaryTableToolStripMenuItem.Text = "Summary Table";
            // 
            // disableSummaryToolStripMenuItem
            // 
            this.disableSummaryToolStripMenuItem.BackColor = System.Drawing.Color.Transparent;
            this.disableSummaryToolStripMenuItem.Name = "disableSummaryToolStripMenuItem";
            this.disableSummaryToolStripMenuItem.Size = new System.Drawing.Size(111, 20);
            this.disableSummaryToolStripMenuItem.Text = "Disable Summary";
            this.disableSummaryToolStripMenuItem.Click += new System.EventHandler(this.disableSummaryToolStripMenuItem_Click);
            // 
            // FilterdataGridView
            // 
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.FilterdataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle7;
            this.FilterdataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.FilterdataGridView.DefaultCellStyle = dataGridViewCellStyle8;
            this.FilterdataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.FilterdataGridView.Location = new System.Drawing.Point(0, 24);
            this.FilterdataGridView.Name = "FilterdataGridView";
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.FilterdataGridView.RowHeadersDefaultCellStyle = dataGridViewCellStyle9;
            this.FilterdataGridView.Size = new System.Drawing.Size(453, 172);
            this.FilterdataGridView.TabIndex = 0;
            // 
            // menuStrip3
            // 
            this.menuStrip3.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.filterDataToolStripMenuItem});
            this.menuStrip3.Location = new System.Drawing.Point(0, 0);
            this.menuStrip3.Name = "menuStrip3";
            this.menuStrip3.Size = new System.Drawing.Size(453, 24);
            this.menuStrip3.TabIndex = 1;
            this.menuStrip3.Text = "menuStrip3";
            // 
            // filterDataToolStripMenuItem
            // 
            this.filterDataToolStripMenuItem.Enabled = false;
            this.filterDataToolStripMenuItem.Name = "filterDataToolStripMenuItem";
            this.filterDataToolStripMenuItem.Size = new System.Drawing.Size(72, 20);
            this.filterDataToolStripMenuItem.Text = "Filter Data";
            // 
            // menuStrip4
            // 
            this.menuStrip4.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.chartSummaryToolStripMenuItem,
            this.disableChartToolStripMenuItem,
            this.disableChartSummaryToolStripMenuItem});
            this.menuStrip4.Location = new System.Drawing.Point(0, 0);
            this.menuStrip4.Name = "menuStrip4";
            this.menuStrip4.Size = new System.Drawing.Size(1262, 24);
            this.menuStrip4.TabIndex = 1;
            this.menuStrip4.Text = "menuStrip4";
            // 
            // chartSummaryToolStripMenuItem
            // 
            this.chartSummaryToolStripMenuItem.Enabled = false;
            this.chartSummaryToolStripMenuItem.Name = "chartSummaryToolStripMenuItem";
            this.chartSummaryToolStripMenuItem.Size = new System.Drawing.Size(102, 20);
            this.chartSummaryToolStripMenuItem.Text = "Chart Summary";
            // 
            // disableChartToolStripMenuItem
            // 
            this.disableChartToolStripMenuItem.Name = "disableChartToolStripMenuItem";
            this.disableChartToolStripMenuItem.Size = new System.Drawing.Size(89, 20);
            this.disableChartToolStripMenuItem.Text = "Disable Chart";
            this.disableChartToolStripMenuItem.Click += new System.EventHandler(this.disableChartToolStripMenuItem_Click);
            // 
            // disableChartSummaryToolStripMenuItem
            // 
            this.disableChartSummaryToolStripMenuItem.Name = "disableChartSummaryToolStripMenuItem";
            this.disableChartSummaryToolStripMenuItem.Size = new System.Drawing.Size(143, 20);
            this.disableChartSummaryToolStripMenuItem.Text = "Disable Chart Summary";
            this.disableChartSummaryToolStripMenuItem.Click += new System.EventHandler(this.disableChartSummaryToolStripMenuItem_Click);
            // 
            // splitContainer7
            // 
            this.splitContainer7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer7.Location = new System.Drawing.Point(0, 0);
            this.splitContainer7.Name = "splitContainer7";
            // 
            // splitContainer7.Panel1
            // 
            this.splitContainer7.Panel1.Controls.Add(this.BpdataGridView);
            this.splitContainer7.Panel1.Controls.Add(this.menuStrip8);
            // 
            // splitContainer7.Panel2
            // 
            this.splitContainer7.Panel2.Controls.Add(this.label6);
            this.splitContainer7.Panel2.Controls.Add(this.textBox1);
            this.splitContainer7.Panel2.Controls.Add(this.chart1);
            this.splitContainer7.Size = new System.Drawing.Size(1262, 275);
            this.splitContainer7.SplitterDistance = 490;
            this.splitContainer7.TabIndex = 6;
            // 
            // BpdataGridView
            // 
            dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.BpdataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle10;
            this.BpdataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.BpdataGridView.DefaultCellStyle = dataGridViewCellStyle11;
            this.BpdataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BpdataGridView.Location = new System.Drawing.Point(0, 24);
            this.BpdataGridView.Name = "BpdataGridView";
            dataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.BpdataGridView.RowHeadersDefaultCellStyle = dataGridViewCellStyle12;
            this.BpdataGridView.Size = new System.Drawing.Size(490, 251);
            this.BpdataGridView.TabIndex = 0;
            // 
            // menuStrip8
            // 
            this.menuStrip8.Location = new System.Drawing.Point(0, 0);
            this.menuStrip8.Name = "menuStrip8";
            this.menuStrip8.Size = new System.Drawing.Size(490, 24);
            this.menuStrip8.TabIndex = 1;
            this.menuStrip8.Text = "menuStrip8";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(732, 34);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(401, 13);
            this.label6.TabIndex = 7;
            this.label6.Text = "Eg. Voltage1 = \'2.375_TTTT\' or Voltage1 = \'1.71_TTTT\' or Voltage1 = \'3.63_TTTT\'";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(735, 50);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(209, 20);
            this.textBox1.TabIndex = 6;
            this.textBox1.Text = "Voltage1 = \'2.375_TTTT\' or Voltage1 = \'1.71_TTTT\' or Voltage1 = \'3.63_TTTT\'";
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageSummaryTool);
            this.tabControl1.Controls.Add(this.tabPageEdit);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 24);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1276, 696);
            this.tabControl1.TabIndex = 7;
            // 
            // tabPageSummaryTool
            // 
            this.tabPageSummaryTool.Controls.Add(this.splitContainer1);
            this.tabPageSummaryTool.Location = new System.Drawing.Point(4, 22);
            this.tabPageSummaryTool.Name = "tabPageSummaryTool";
            this.tabPageSummaryTool.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSummaryTool.Size = new System.Drawing.Size(1268, 670);
            this.tabPageSummaryTool.TabIndex = 1;
            this.tabPageSummaryTool.Text = "Summary Tool";
            this.tabPageSummaryTool.UseVisualStyleBackColor = true;
            // 
            // tabPageEdit
            // 
            this.tabPageEdit.Controls.Add(this.EditdataGridView);
            this.tabPageEdit.Location = new System.Drawing.Point(4, 22);
            this.tabPageEdit.Name = "tabPageEdit";
            this.tabPageEdit.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageEdit.Size = new System.Drawing.Size(1268, 670);
            this.tabPageEdit.TabIndex = 0;
            this.tabPageEdit.Text = "Edit";
            this.tabPageEdit.UseVisualStyleBackColor = true;
            // 
            // EditdataGridView
            // 
            dataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle13.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle13.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle13.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle13.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.EditdataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle13;
            this.EditdataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle14.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle14.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle14.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle14.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle14.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.EditdataGridView.DefaultCellStyle = dataGridViewCellStyle14;
            this.EditdataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.EditdataGridView.Location = new System.Drawing.Point(3, 3);
            this.EditdataGridView.Name = "EditdataGridView";
            dataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle15.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle15.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle15.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle15.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle15.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.EditdataGridView.RowHeadersDefaultCellStyle = dataGridViewCellStyle15;
            this.EditdataGridView.Size = new System.Drawing.Size(1262, 664);
            this.EditdataGridView.TabIndex = 0;
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "notifyIcon1";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.Click += new System.EventHandler(this.notifyIcon1_Click);
            // 
            // SummaryToolFrm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(1276, 742);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "SummaryToolFrm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Summary Tool";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SummaryToolForm_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.SumdataGridView)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer3.Panel1.ResumeLayout(false);
            this.splitContainer3.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer3)).EndInit();
            this.splitContainer3.ResumeLayout(false);
            this.splitContainer4.Panel1.ResumeLayout(false);
            this.splitContainer4.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer4)).EndInit();
            this.splitContainer4.ResumeLayout(false);
            this.splitContainer5.Panel1.ResumeLayout(false);
            this.splitContainer5.Panel1.PerformLayout();
            this.splitContainer5.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer5)).EndInit();
            this.splitContainer5.ResumeLayout(false);
            this.menuStrip9.ResumeLayout(false);
            this.menuStrip9.PerformLayout();
            this.splitContainer11.Panel1.ResumeLayout(false);
            this.splitContainer11.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer11)).EndInit();
            this.splitContainer11.ResumeLayout(false);
            this.splitContainer12.Panel1.ResumeLayout(false);
            this.splitContainer12.Panel1.PerformLayout();
            this.splitContainer12.Panel2.ResumeLayout(false);
            this.splitContainer12.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer12)).EndInit();
            this.splitContainer12.ResumeLayout(false);
            this.menuStrip12.ResumeLayout(false);
            this.menuStrip12.PerformLayout();
            this.menuStrip14.ResumeLayout(false);
            this.menuStrip14.PerformLayout();
            this.splitContainer13.Panel1.ResumeLayout(false);
            this.splitContainer13.Panel1.PerformLayout();
            this.splitContainer13.Panel2.ResumeLayout(false);
            this.splitContainer13.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer13)).EndInit();
            this.splitContainer13.ResumeLayout(false);
            this.menuStrip7.ResumeLayout(false);
            this.menuStrip7.PerformLayout();
            this.menuStrip13.ResumeLayout(false);
            this.menuStrip13.PerformLayout();
            this.splitContainer6.Panel1.ResumeLayout(false);
            this.splitContainer6.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer6)).EndInit();
            this.splitContainer6.ResumeLayout(false);
            this.splitContainer9.Panel1.ResumeLayout(false);
            this.splitContainer9.Panel1.PerformLayout();
            this.splitContainer9.Panel2.ResumeLayout(false);
            this.splitContainer9.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer9)).EndInit();
            this.splitContainer9.ResumeLayout(false);
            this.menuStrip6.ResumeLayout(false);
            this.menuStrip6.PerformLayout();
            this.menuStrip10.ResumeLayout(false);
            this.menuStrip10.PerformLayout();
            this.splitContainer10.Panel1.ResumeLayout(false);
            this.splitContainer10.Panel1.PerformLayout();
            this.splitContainer10.Panel2.ResumeLayout(false);
            this.splitContainer10.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer10)).EndInit();
            this.splitContainer10.ResumeLayout(false);
            this.menuStrip5.ResumeLayout(false);
            this.menuStrip5.PerformLayout();
            this.menuStrip11.ResumeLayout(false);
            this.menuStrip11.PerformLayout();
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.splitContainer8.Panel1.ResumeLayout(false);
            this.splitContainer8.Panel1.PerformLayout();
            this.splitContainer8.Panel2.ResumeLayout(false);
            this.splitContainer8.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer8)).EndInit();
            this.splitContainer8.ResumeLayout(false);
            this.splitContainer14.Panel1.ResumeLayout(false);
            this.splitContainer14.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer14)).EndInit();
            this.splitContainer14.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.HighSummaryDataGridView)).EndInit();
            this.menuStrip2.ResumeLayout(false);
            this.menuStrip2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.FilterdataGridView)).EndInit();
            this.menuStrip3.ResumeLayout(false);
            this.menuStrip3.PerformLayout();
            this.menuStrip4.ResumeLayout(false);
            this.menuStrip4.PerformLayout();
            this.splitContainer7.Panel1.ResumeLayout(false);
            this.splitContainer7.Panel1.PerformLayout();
            this.splitContainer7.Panel2.ResumeLayout(false);
            this.splitContainer7.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer7)).EndInit();
            this.splitContainer7.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.BpdataGridView)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPageSummaryTool.ResumeLayout(false);
            this.tabPageEdit.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.EditdataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private void selectInputsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DtPlots1.sm.showInput(); //DtPlots1.summaryInit();
            SummaryPlotInit();
            toolStripStatusLabel1.Text = DtPlots1.toolStripStatusLabel;
        }
        public void SummaryPlotInit()
        {
            AllEntryDNDLV.Items.Clear();

            InputRowDNDLV.Items.Clear();
            InputColumnDNDLV.Items.Clear();
            YAxisOutputDNDLV.Items.Clear();
            YAxisCombinedOutputDNDLV.Items.Clear();
            XaxisRowDNDLV.Items.Clear();
            XaxisColumnDNDLV.Items.Clear();
            ChartSeriesRowDNDLV.Items.Clear();
            ChartSeriesColumnDNDLV.Items.Clear();
            DtPlots1.summaryInit();

            DataTable dtRaw = DtPlots1.sm.RawData();
            if (dtRaw != null)
            {
                for (int i = 0; i < dtRaw.Columns.Count; i++)
                {
                    AllEntryDNDLV.Items.Add(SummaryMainClass.RestoreData(DtPlots1.sm.dt.Columns[i].ToString()));
                }
                SumdataGridView.DataSource = dtRaw.DefaultView;
            }
            //chart1.ChartAreas.Clear();
            toolStripStatusLabel1.Text = DtPlots1.toolStripStatusLabel;
        }

        private void SummaryToolForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            FormSerializor.FormSerilizor.Serialise(this);
        }

        private void restoreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            restoreData();
        }

        public void restoreData()
        {
            InitializeInputFields();
            DtPlots1 = new DtPlots(InputFields1, new DataGridView[] { SumdataGridView, HighSummaryDataGridView, FilterdataGridView, BpdataGridView, EditdataGridView }, chart1);

            DtPlots1.summaryInit();
            SummaryPlotInit();
            toolStripStatusLabel1.Text = DtPlots1.toolStripStatusLabel;
        }
        public void export2PPT(string sel)
        {
            string filename = System.IO.Path.GetFileName(DtPlots1.sm.inputSetting.DataPath);
            string path = "";
            if (sel == "CHART")
            {
                path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(DtPlots1.sm.inputSetting.DataPath), System.IO.Path.GetFileNameWithoutExtension(DtPlots1.sm.inputSetting.DataPath) + "ChartReport.ppt");
                PptLayer.ShowPresentation(path);
                //Generate Save Chart to ppt
                DtPlots1.GenerateNSaveChart(path);
            }
            if (sel == "SUMMARY")
            {
                path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(DtPlots1.sm.inputSetting.DataPath), System.IO.Path.GetFileNameWithoutExtension(DtPlots1.sm.inputSetting.DataPath) + "SummaryReport.ppt");
                PptLayer.ShowPresentation(path);
                DtPlots1.SaveSummary(path);
            }
            PptLayer.SavePresentation(path);
            MessageBox.Show("Report generation Completed !!!");
        }

        private void SumdataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            toolStripStatusLabel1.Text = "Retriving the data...";
            statusStrip1.Refresh();
            this.Refresh();
            if (e.RowIndex != -1 & e.ColumnIndex != -1) DtPlots1.refreshDgv(2, e);
            toolStripStatusLabel1.Text = "Ready";
        }

        private void table1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(DtPlots1.sm.inputSetting.DataPath), System.IO.Path.GetFileNameWithoutExtension(DtPlots1.sm.inputSetting.DataPath) + "Summary.csv");
            if (File.Exists(path))
            {
                if (IsFileinUse(path)) { MessageBox.Show("Please Close the file and try again: " + path); return; }

                File.Delete(path);
            }
            DtPlots1.GenerateNSaveSummary(path, true, "CSV");
        }

        private bool IsFileinUse(string filename)
        {
            FileInfo file = new FileInfo(filename);
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
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
            return false;
        }

        private void table2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportToCsv.ToCsV(FilterdataGridView, System.IO.Path.Combine(System.IO.Path.GetDirectoryName(DtPlots1.sm.inputSetting.DataPath), "FilteredData.csv"));
        }
        private void chartSummaryToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ExportToCsv.ToCsV(BpdataGridView, System.IO.Path.Combine(System.IO.Path.GetDirectoryName(DtPlots1.sm.inputSetting.DataPath), System.IO.Path.GetFileNameWithoutExtension(DtPlots1.sm.inputSetting.DataPath) + "_ChartSummary.csv"));
        }
        private void highLevelSummaryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportToCsv.ToCsV(HighSummaryDataGridView
                , System.IO.Path.Combine(System.IO.Path.GetDirectoryName(DtPlots1.sm.inputSetting.DataPath), System.IO.Path.GetFileNameWithoutExtension(DtPlots1.sm.inputSetting.DataPath) + "_HighLevelSummary.csv"));
        }

        public void chart1_Customize(object sender, EventArgs e)
        {
            DtPlots1.ChartRedraw();
            if (DtPlots1.toolStripStatusLabel == null && toolStripStatusLabel1.Text == "") toolStripStatusLabel1.Text = "Ready";
        }

        private void chart1_PostPaint(object sender, ChartPaintEventArgs e)
        {
            if (e.ChartElement is Chart)
            {
                // PostPaint(e);
            }
        }
        public void PostPaint(ChartPaintEventArgs e)
        {
            // create text to draw
            String TextToDraw;
            TextToDraw = "";
            TextToDraw += "";

            // get graphics tools
            Graphics g = e.ChartGraphics.Graphics;
            Font DrawFont = System.Drawing.SystemFonts.CaptionFont;
            Brush DrawBrush = Brushes.Black;

            // see how big the text will be
            int TxtWidth = (int)g.MeasureString(TextToDraw, DrawFont).Width;
            int TxtHeight = (int)g.MeasureString(TextToDraw, DrawFont).Height;

            // where to draw
            int x = 15;  // a few pixels from the left border

            int y = (int)e.Chart.Height;
            y = y - TxtHeight - 5; // a few pixels off the bottom

            // draw the string        
            g.DrawString(TextToDraw, DrawFont, DrawBrush, x, y);
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void createSpecToolStripMenuItem_Click(object sender, EventArgs e)
        {
            createSpec();
        }
        public void createSpec()
        {
            toolStripStatusLabel1.Text = "Creating the spec";
            //Referesh
            DtPlots1.refreshDgv();
            DtPlots1.CreateSpec();
            saveSpec();
            toolStripStatusLabel1.Text = DtPlots1.toolStripStatusLabel;
            if (DtPlots1.toolStripStatusLabel == null) toolStripStatusLabel1.Text = "Spec created, saved in data folder.";
        }
        public void saveSpec()
        {
            ExportToCsv.ToCsV(EditdataGridView, System.IO.Path.Combine(System.IO.Path.GetDirectoryName(DtPlots1.sm.inputSetting.DataPath), System.IO.Path.GetFileNameWithoutExtension(DtPlots1.sm.inputSetting.DataPath) + "_RecommandedSpec.csv"));
        }

        private void NSigmatextBox_TextChanged(object sender, EventArgs e)
        {
            Double.TryParse(NSigmatextBox.Text, out DtPlots1.InputFields1.NSigma);
            refreshDgv();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            DtPlots1.TypicalLotFilter = textBox1.Text;
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            refreshDgv();
        }
        private Keys prevKey;
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (prevKey == Keys.D && keyData == Keys.S) { disableSummary(); return true; }
            if (prevKey == Keys.D && keyData == Keys.C) { disableChart(); return true; }
            if (prevKey == Keys.D && keyData == Keys.U) { disableChartSummary(); return true; }
            if (prevKey == Keys.D && keyData == Keys.I) { disableInputSection(); return true; }
            if (keyData == Keys.F5) refreshDgv();
            if (keyData == Keys.I) { DtPlots1.sm.showInput(); SummaryPlotInit(); }
            if (keyData == Keys.R) restoreData();
            if (prevKey == Keys.P && keyData == Keys.C) export2PPT("CHART");
            if (prevKey == Keys.P && keyData == Keys.S) export2PPT("SUMMARY");
            if (keyData == Keys.S) createSpec();
            if (keyData == Keys.N) saveSpec();
            if (keyData == Keys.X) this.Close();

            prevKey = keyData;
            return base.ProcessCmdKey(ref msg, keyData);

        }

        private void releaseToQAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Updater.ReleaseEng("Engineering");
        }

        private void AdminMode_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you want to enable to Admin mode?", "Change to Admin Mode", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                releaseToQAToolStripMenuItem.Visible = true;
                releaseProductionToolStripMenuItem.Visible = true;
            }

        }

        private void updateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //To get the location the assembly normally resides on disk or the install directory
            string path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;

            //once you have the path you get the directory with:
            var directory = System.IO.Path.GetDirectoryName(path);
            System.Diagnostics.Process.Start(System.IO.Path.Combine(directory, "Update.exe"));
            Close();
        }

        private void disableSummaryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            disableSummary();
        }

        private void disableSummaryToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            disableSummary();
        }

        private void disableSummary()
        {
            if (disableSummaryToolStripMenuItem.Text == "Disable Summary")
            {//Disable Summary
                disableSummaryToolStripMenuItem1.Text = "Enable Summary              DS";
                disableSummaryToolStripMenuItem.Text = "Enable Summary";
                splitContainer2.Panel1Collapsed = true; splitContainer2.Panel1.Hide();
                if (disableChartToolStripMenuItem.Text == "Enable Chart" && disableSummaryToolStripMenuItem.Text == "Enable Summary")
                { splitContainer1.Panel2Collapsed = true; splitContainer1.Panel2.Hide(); splitContainer2.Visible = false; }
                return;
            }
            else
            {//Enable summary
                splitContainer1.Visible = true;
                splitContainer2.Visible = true;
                splitContainer1.Panel2Collapsed = false; splitContainer1.Panel2.Show();
                splitContainer2.Panel1Collapsed = false; splitContainer2.Panel1.Show();
                disableSummaryToolStripMenuItem1.Text = "Disable Summary             DS";
                disableSummaryToolStripMenuItem.Text = "Disable Summary";
            }
        }
        private void disableChartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            disableChart();
        }

        private void disableChartDCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            disableChart();
        }
        private void disableChart()
        {
            if (disableChartToolStripMenuItem.Text == "Enable Chart")
            { //Enable chart
                DtPlots1.enableChartFlag = true;
                splitContainer1.Visible = true;
                splitContainer2.Visible = true;
                splitContainer7.Visible = true;
                splitContainer1.Panel2Collapsed = false; splitContainer1.Panel2.Show();
                splitContainer2.Panel2Collapsed = false; splitContainer2.Panel2.Show();
                splitContainer7.Panel2Collapsed = false; splitContainer7.Panel2.Show();
                disableChartDCToolStripMenuItem.Text = "Diable Chart                DC";
                disableChartToolStripMenuItem.Text = "Disable Chart"; return;
            }
            else
            {//Disable chart
                DtPlots1.enableChartFlag = false;
                disableChartDCToolStripMenuItem.Text = "Enable Chart                DC";
                disableChartToolStripMenuItem.Text = "Enable Chart";
                splitContainer7.Panel2Collapsed = true; splitContainer7.Panel2.Hide();
                if (disableChartSummaryToolStripMenuItem.Text == "Enable Chart Summary")
                {//disable the container as chart summary is also closed
                    splitContainer7.Visible = false;
                    splitContainer2.Panel2Collapsed = true; splitContainer2.Panel2.Hide();
                    if (disableSummaryToolStripMenuItem.Text == "Enable Summary")
                    { //close the container as Summary is also closed
                        splitContainer2.Visible = false;
                        splitContainer1.Panel2Collapsed = true; splitContainer1.Panel2.Hide();
                        if (disableInputSectionToolStripMenuItem1.Text == "Enable Input Section") //close all container
                            splitContainer1.Visible = false;
                    }
                }
            }
        }

        private void disableChartSummaryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            disableChartSummary();
        }

        private void disableChartSummaryDUToolStripMenuItem_Click(object sender, EventArgs e)
        {
            disableChartSummary();
        }
        private void disableChartSummary()
        {
            if (disableChartSummaryToolStripMenuItem.Text == "Enable Chart Summary")
            {//Enable chart summary
                DtPlots1.enableChartSumFlag = true;

                splitContainer1.Visible = true;
                splitContainer2.Visible = true;
                splitContainer7.Visible = true;
                splitContainer1.Panel2Collapsed = false; splitContainer1.Panel2.Show();
                splitContainer2.Panel2Collapsed = false; splitContainer2.Panel2.Show();
                splitContainer7.Panel1Collapsed = false; splitContainer7.Panel1.Show();
                disableChartSummaryDUToolStripMenuItem.Text = "Disable Chart Summary       DU";
                disableChartSummaryToolStripMenuItem.Text = "Disable Chart Summary"; return;
            }
            else
            {//disable chart summary
                DtPlots1.enableChartSumFlag = false;

                disableChartSummaryDUToolStripMenuItem.Text = "Enable Chart Summary        DU";
                disableChartSummaryToolStripMenuItem.Text = "Enable Chart Summary";
                splitContainer7.Panel1Collapsed = true; splitContainer7.Panel1.Hide();
                if (disableChartToolStripMenuItem.Text == "Enable Chart")
                {//close the container as chart is also closed
                    splitContainer7.Visible = false;
                    splitContainer2.Panel2Collapsed = true; splitContainer2.Panel2.Hide();
                    if (disableSummaryToolStripMenuItem.Text == "Enable Summary")
                    {//close the container as Summary is also closed
                        splitContainer2.Visible = false;
                        splitContainer1.Panel2Collapsed = true; splitContainer1.Panel2.Hide();
                        if (disableInputSectionToolStripMenuItem1.Text == "Enable Input Section") //close all container
                            splitContainer1.Visible = false;
                    }
                }
            }
        }
        private void disableInputSectionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            disableInputSection();
        }
        private void disableInputSection()
        {
            if (disableInputSectionToolStripMenuItem1.Text == "Enable Input Section")
            {//Enable Input
                splitContainer1.Visible = true;
                disableInputSectionToolStripMenuItem1.Text = "Disable Input Section";
                disableInputSectionToolStripMenuItem.Text = "Disable Input Section       DI";
                splitContainer1.Panel1Collapsed = false; splitContainer1.Panel1.Show();
            }
            else
            {//Disable Input
                disableInputSectionToolStripMenuItem1.Text = "Enable Input Section";
                disableInputSectionToolStripMenuItem.Text = "Enable Input Section       DI";
                splitContainer1.Panel1Collapsed = true; splitContainer1.Panel1.Hide();
                if (disableChartToolStripMenuItem.Text == "Enable Chart" && disableChartSummaryToolStripMenuItem.Text == "Enable Chart Summary" && disableSummaryToolStripMenuItem.Text == "Enable Summary")
                {//close the container as chart is also closed
                    splitContainer1.Visible = false;
                }
            }
        }

        private void disableInputSectionToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            disableInputSection();
        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void minimizeToTrayToolStripMenuItem_Click(object sender, EventArgs e)
        {
            notifyIcon1.BalloonTipTitle = "Char Tool";
            notifyIcon1.BalloonTipText = "Your App is minimized here";
            notifyIcon1.Visible = true;
            notifyIcon1.ShowBalloonTip(5000);
            this.Hide();
        }

        private void summaryPSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            export2PPT("SUMMARY");
            toolStripStatusLabel1.Text = DtPlots1.toolStripStatusLabel;
        }

        private void chartSummaryPCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (disableChartToolStripMenuItem.Text == "Disable Chart") export2PPT("CHART");
            if (disableChartToolStripMenuItem.Text == "Disable Chart Summary") export2PPT("SUMMARY");
            toolStripStatusLabel1.Text = DtPlots1.toolStripStatusLabel;
        }

        /// <summary>
        /// Refreshes the Datagridview, reloads the input fields from dropdown boxes
        /// </summary>
        /// <returns></returns>
        public bool refreshDgv()
        {
            if (InitializeInputFields())
            {
                DtPlots1.InputFields1 = InputFields1;
                DtPlots1.refreshDgv();
                if (DtPlots1.toolStripStatusLabel != null) toolStripStatusLabel1.Text = DtPlots1.toolStripStatusLabel;
            }
            return true;
        }

        private void releaseProductionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Updater.ReleaseProduction("Production");
        }

        private void revisionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string About = string.Format(System.Globalization.CultureInfo.InvariantCulture, @"YourApp Version {0}.{1}.{2} (r{3})", v.Major, v.Minor, v.Build, v.Revision);

            var entryAssembly = System.Reflection.Assembly.GetEntryAssembly();
            var fileInfo = new FileInfo(entryAssembly.Location);
            var buildDate = fileInfo.LastWriteTime;

            MessageBox.Show(About + " Built on " + buildDate, "Revision");
        }

        private void InputRowDNDLV_DragDrop(object sender, DragEventArgs e)
        {
            refreshDgv();
        }

        private void InputColumnDNDLV_DragDrop(object sender, DragEventArgs e)
        {
            refreshDgv();
        }

        private void XaxisRowDNDLV_DragDrop(object sender, DragEventArgs e)
        {
            refreshDgv();
        }

        private void XaxisColumnDNDLV_DragDrop(object sender, DragEventArgs e)
        {
            refreshDgv();
        }

        private void ChartSeriesRowDNDLV_DragDrop(object sender, DragEventArgs e)
        {
            refreshDgv();
        }

        private void ChartSeriesColumnDNDLV_DragDrop(object sender, DragEventArgs e)
        {
            refreshDgv();
        }

        private void YAxisOutputDNDLV_DragDrop(object sender, DragEventArgs e)
        {
            refreshDgv();
        }

        private void YAxisCombinedOutputDNDLV_DragDrop(object sender, DragEventArgs e)
        {
            refreshDgv();
        }

        private void YAxisOutputDNDLV_DragDrop_1(object sender, DragEventArgs e)
        {
            refreshDgv();
        }

        private void StatisticsSelectedDNDLV_DragDrop(object sender, DragEventArgs e)
        {
            refreshDgv();
        }

        private void StatisticsDNDLV_DragDrop(object sender, DragEventArgs e)
        {
            refreshDgv();
        }

        private void AllEntryDNDLV_DragDrop(object sender, DragEventArgs e)
        {
            refreshDgv();
        }
        private void showPopup(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                DtPlots1.f = new PopupForm();
                //Subscribe to event
                DtPlots1.f.NewDragNDropChanged += new EventHandler<TextEventArgs>(PopupForm_NewTextChanged);

                DtPlots1.ShowPopup(((ListView)sender).SelectedItems[0].Text.ToString());

                //Unsubscribe from event
                DtPlots1.f.NewDragNDropChanged -= PopupForm_NewTextChanged;
                DtPlots1.f.Dispose();
                DtPlots1.f = null;

                return;
            }

        }
        private void InputRowDNDLV_MouseClick(object sender, MouseEventArgs e)
        {
            showPopup(sender, e);
            refreshDgv();
        }

        // Our new event (form2.NewTextChanged) handler
        private void PopupForm_NewTextChanged(object sender, TextEventArgs e)
        {
            //Text = e.Text;
        }

        private void InputColumnDNDLV_MouseClick(object sender, MouseEventArgs e)
        {
            showPopup(sender, e);
            refreshDgv();
        }

        private void XaxisRowDNDLV_MouseClick(object sender, MouseEventArgs e)
        {
            showPopup(sender, e);
            refreshDgv();
        }

        private void XaxisColumnDNDLV_MouseClick(object sender, MouseEventArgs e)
        {
            showPopup(sender, e);
            refreshDgv();
        }

        private void ChartSeriesRowDNDLV_MouseClick(object sender, MouseEventArgs e)
        {
            showPopup(sender, e);
            refreshDgv();
        }

        private void ChartSeriesColumnDNDLV_MouseClick(object sender, MouseEventArgs e)
        {
            showPopup(sender, e);
            refreshDgv();
        }

        private void AllEntryDNDLV_MouseClick(object sender, MouseEventArgs e)
        {
            showPopup(sender, e);
            refreshDgv();
        }

    }
}
