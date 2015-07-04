namespace SummaryToolForm
{
    partial class PopupForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
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
            this.DragNDropListView1 = new DragNDrop.DragNDropListView();
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader6 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SuspendLayout();
            // 
            // DragNDropListView1
            // 
            this.DragNDropListView1.AllowDrop = true;
            this.DragNDropListView1.AllowReorder = true;
            this.DragNDropListView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader5,
            this.columnHeader6});
            this.DragNDropListView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DragNDropListView1.FullRowSelect = true;
            this.DragNDropListView1.LineColor = System.Drawing.Color.Olive;
            this.DragNDropListView1.Location = new System.Drawing.Point(0, 0);
            this.DragNDropListView1.Name = "DragNDropListView1";
            this.DragNDropListView1.Size = new System.Drawing.Size(284, 262);
            this.DragNDropListView1.TabIndex = 2;
            this.DragNDropListView1.UseCompatibleStateImageBehavior = false;
            this.DragNDropListView1.View = System.Windows.Forms.View.List;
            this.DragNDropListView1.Click += new System.EventHandler(this.DragNDropListView1_Click);
            // 
            // columnHeader5
            // 
            this.columnHeader5.Width = 190;
            // 
            // columnHeader6
            // 
            this.columnHeader6.Width = 244;
            // 
            // PopupForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.DragNDropListView1);
            this.MinimizeBox = false;
            this.Name = "PopupForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Text = "Choose input entries";
            this.TransparencyKey = System.Drawing.Color.DimGray;
            this.Deactivate += new System.EventHandler(this.PopupForm_Deactivate_1);
            this.ResumeLayout(false);

        }

        #endregion

        public DragNDrop.DragNDropListView DragNDropListView1;
        private System.Windows.Forms.ColumnHeader columnHeader5;
        private System.Windows.Forms.ColumnHeader columnHeader6;

    }
}