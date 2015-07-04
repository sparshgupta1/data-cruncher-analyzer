using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Update
{
    public partial class InstallerForm : Form
    {
        public InstallerForm(string app)
        {
            InitializeComponent();

            string LocalDir = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            if (!System.IO.Directory.Exists(LocalDir)) System.IO.Directory.CreateDirectory(LocalDir);
            Updater.DirectoryCopy(Updater.ServerDir, System.IO.Path.Combine(LocalDir, "Summary Tool\\Software"));
            Updater.CreateShortcut(System.IO.Path.Combine(LocalDir, app));
        }
    }
}
