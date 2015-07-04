using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using XmlSerialization;
using System.Windows.Forms;

namespace Update
{
    public class Updater
    {
        public Settings inputSetting; // = new settings();
        static object xobj = new Settings();
        public Serializer sl = new Serializer(ref xobj);
        public static string ServerDir;
        public static string FolderStructure = @"\SummaryTool\Software\";
        public static string Production = "Production";
        public static string Engineering = "Engineering";
        public Updater()
        {

            inputSetting = (Settings)xobj;
            //Kill the process
            Process[] processes = Process.GetProcessesByName("SummaryToolForm");
            if (processes.Length > 0)
                processes[0].Kill();

            // Get current time zone.
            TimeZone zone = TimeZone.CurrentTimeZone;
            string standard = zone.StandardName;
            ServerDir = inputSetting.ITCDirectory;
            if (standard != "India Standard Time") { ServerDir = inputSetting.AustinDirectory; }

            string LocalDir = System.IO.Directory.GetCurrentDirectory();
            
            //Check the program is running as updater or Installer
            if (System.IO.Path.GetFileNameWithoutExtension(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName) == "Update")
            {
                //Skip the development copy
                if (!System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName).Contains("bin\\Debug")) CopyDir(ServerDir, LocalDir);
            }
            else //runs as installer
            {
                InstallerForm f = new InstallerForm(FolderStructure + Production + "\\Summary Tool.exe");
                LocalDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),"Summary Tool");
                if (!System.IO.Directory.Exists(LocalDir)) System.IO.Directory.CreateDirectory(LocalDir);
                DirectoryCopy(ServerDir, LocalDir);
                appShortcutToDesktop("test"); //***cb*******
            }
        }
        /// <summary>
        /// Releases the software to Engineering mode
        /// </summary>
        /// <returns></returns>
        public bool ReleaseEng()
        {
            Settings inputSetting = (Settings)xobj;

            //Release the bin to Eng 
            string LocalPath = Directory.GetCurrentDirectory();

            DirectoryCopy(LocalPath, System.IO.Path.Combine(inputSetting.ITCDirectory, "Engineering")); //Copy the bin files to ITC server

            DirectoryCopy(LocalPath, System.IO.Path.Combine(inputSetting.AustinDirectory, "Engineering")); //Copy the bin files to Austin servers.

            return true;
        }
        /// <summary>
        /// Releases the software to production mode
        /// </summary>
        /// <returns></returns>
        public bool ReleaseProduction()
        {
            Settings inputSetting = (Settings)xobj;

            //Release the bin to Production
            if (!(MessageBox.Show("Have you validated Engineering release?", "Important Query", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes)) return false;
            
            string LocalPath = Directory.GetCurrentDirectory();

            DirectoryCopy(LocalPath, System.IO.Path.Combine(inputSetting.ITCDirectory, "Production")); //Copy the bin files to ITC server

            DirectoryCopy(LocalPath, System.IO.Path.Combine(inputSetting.AustinDirectory, "Production")); //Copy the bin files to Austin servers.

            return true;
        }        
        /// <summary>
        /// Copies the content of the directory. If the file is in use it quietly skips
        /// </summary>
        /// <param name="sourceDir"></param>
        /// <param name="DestDir"></param>
        public static void CopyDir(string sourceDir, string DestDir)
        {
            if (!System.IO.Directory.Exists(DestDir)) System.IO.Directory.CreateDirectory(DestDir);
            foreach (var file in Directory.GetFiles(sourceDir))
            {
                try
                {
                    string fPath = Path.Combine(DestDir, Path.GetFileName(file));
                    if (System.IO.File.Exists(fPath)) File.Delete(fPath);
                    File.Copy(file, fPath);
                }
                catch { }
            }

            foreach (var directory in Directory.GetDirectories(sourceDir))
            {
                string tempPath = Path.Combine(DestDir, Path.GetFileName(directory));
                try { if (System.IO.Directory.Exists(tempPath)) Directory.Delete(tempPath, true); }
                catch { }
                Directory.CreateDirectory(tempPath);
                foreach (var file in Directory.GetFiles(directory))
                    File.Copy(file, Path.Combine(tempPath, Path.GetFileName(file)));
            }
        }
        public static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs=true)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            // If the destination directory doesn't exist, create it. 
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath,true);
            }

            // If copying subdirectories, copy them and their contents to new location. 
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }
        public static void appShortcutToDesktop(string linkName)
        {
            string deskDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            using (StreamWriter writer = new StreamWriter(deskDir + "\\" + Path.GetFileName(linkName) + ".url"))
            {
                string app = System.Reflection.Assembly.GetExecutingAssembly().Location;
                writer.WriteLine("[InternetShortcut]");
                writer.WriteLine("URL=file:///" + app);
                writer.WriteLine("IconIndex=0");
                string icon = app.Replace('\\', '/');
                writer.WriteLine("IconFile=" + icon);
                writer.Flush();
            }
        }

    }
    public class Settings
    {
        public string ITCDirectory = Path.Combine(@"\\flrblr001\windows_data\From-SLI-Fileserver\E\Product Engineering\CHAR DATA\", Updater.FolderStructure);
        public string AustinDirectory = Path.Combine(@"\\silabs.com\apps\Wireline\Timing\\", Updater.FolderStructure);
        //public string OtherReleasePath = new string() {"somePath"};
    }

}
