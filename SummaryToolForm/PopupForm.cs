using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SummaryToolForm
{
    public partial class PopupForm : Form
    {
        // Our event
        public event EventHandler<TextEventArgs> NewDragNDropChanged;
        /* Private field to hold the value that is passed around.
         * Exposed by the NewText property and passed in the NewTextChanged event. */
        private string dragNDropString;
        public PopupForm()
        {
            InitializeComponent();

            //Set the location of the form
            this.Location = new Point(Cursor.Position.X, Cursor.Position.Y);

            this.StartPosition = FormStartPosition.Manual; //NEW CODE
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape) this.Close();
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void DragNDropListView1_Click(object sender, EventArgs e)
        {
            if ((ModifierKeys == Keys.Control)) return;
            List<string> dragNDropList = new List<string> { };
            foreach (ListViewItem SelectedItem in DragNDropListView1.SelectedItems) dragNDropList.Add(SelectedItem.Text);
            dragNDropString = string.Join("qcommaq", dragNDropList);
            DragNDropString = dragNDropString;
        }
        public string DragNDropString
        {
            get { return dragNDropString; }
            /* Setting this property will raise the event if the value is different.
             * As this property has a public getter you can access it and raise the
             * event from any reference to this class. in this example it is set from
             * the textBox.TextChanged handler above. The setter can be marked as
             * as private if required. */
            set
            {
                dragNDropString = value;
                OnNewDragNDropChanged(new TextEventArgs(dragNDropString));
            }
        }
        // Standard event raising pattern.
        protected virtual void OnNewDragNDropChanged(TextEventArgs e)
        {
            // Create a copy of the event to work with
            EventHandler<TextEventArgs> eh = NewDragNDropChanged;
            /* If there are no subscribers, eh will be null so we need to check
             * to avoid a NullReferenceException. */
            if (eh != null)
                eh(this, e);
        }

        private void PopupForm_Deactivate_1(object sender, EventArgs e)
        {
            this.Close();
        }
    }
    public class TextEventArgs : EventArgs
    {
        // Private field exposed by the Text property
        private string text;

        public TextEventArgs(string text)
        {
            this.text = text;
        }

        // Read only property.
        public string Text
        {
            get { return text; }
        }
    }
}
