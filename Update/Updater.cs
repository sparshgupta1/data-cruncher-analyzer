using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using XmlSerialization;
using System.Windows.Forms;

using IWshRuntimeLibrary;

namespace Update
{
    public class Updater
    {
        public Settings inputSetting; // = new settings();
        static object xobj = new Settings();
        public Serializer sl = new Serializer(ref xobj);
        public static string ServerDir;
        public static string FolderStructure = @"Summary Tool\Software\";
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

            string Base = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;

            string LocalDir = System.IO.Path.GetDirectoryName(Base).Replace("file:\\", "");
            //Copies the files from server
            InstallerForm f = new InstallerForm(FolderStructure + Production + "\\SummaryToolForm.exe");
           //Check the program is running as updater or Installer
            if (System.IO.Path.GetFileNameWithoutExtension(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName) == "Update")
            {
                //Skip the development copy
                if (!System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName).Contains("bin\\Debug"))
                {
                    //DirectoryCopy(ServerDir, LocalDir,true,true);
                    MessageBox.Show("Updated Successfully.!");

                    //To get the location the assembly normally resides on disk or the install directory
                    Base = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;

                    LocalDir = System.IO.Path.GetDirectoryName(Base).Replace("file:\\", "");

                    System.Diagnostics.Process.Start(System.IO.Path.Combine(LocalDir, "SummaryToolForm.exe"));
                }
            }
            else //runs as installer
            {
                MessageBox.Show("Installed, use the shortcut created in your desktop (Ctrl+Alt+C)!");
            }
        }
        /// <summary>
        /// Releases the software to Engineering mode
        /// </summary>
        /// <returns></returns>
        public static bool ReleaseEng(string folderStructure = "")
        {
            Settings inputSetting = (Settings)xobj;

            //Release the bin to Eng 
            string Base = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;

            string LocalDir = System.IO.Path.GetDirectoryName(Base).Replace("file:\\", "");

            DirectoryCopy(LocalDir, System.IO.Path.Combine(inputSetting.ITCDirectory, folderStructure)); //Copy the bin files to ITC server

            DirectoryCopy(LocalDir, System.IO.Path.Combine(inputSetting.AustinDirectory, folderStructure)); //Copy the bin files to Austin servers.

            return true;
        }
        /// <summary>
        /// Releases the software to production mode
        /// </summary>
        /// <returns></returns>
        public static bool ReleaseProduction(string folderStructure = "")
        {
            Settings inputSetting = (Settings)xobj;

            //Release the bin to Production
            if (!(MessageBox.Show("Have you validated Engineering release?", "Important Query", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes)) return false;

            string Base = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;

            string LocalDir = System.IO.Path.GetDirectoryName(Base).Replace("file:\\", "");

            DirectoryCopy(LocalDir, System.IO.Path.Combine(inputSetting.ITCDirectory, folderStructure)); //Copy the bin files to ITC server

            DirectoryCopy(LocalDir, System.IO.Path.Combine(inputSetting.AustinDirectory, folderStructure)); //Copy the bin files to Austin servers.

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
                    if (System.IO.File.Exists(fPath)) System.IO.File.Delete(fPath);
                    System.IO.File.Copy(file, fPath);
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }

            foreach (var directory in Directory.GetDirectories(sourceDir))
            {
                string tempPath = Path.Combine(DestDir, Path.GetFileName(directory));
                try { if (System.IO.Directory.Exists(tempPath)) Directory.Delete(tempPath, true); }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                Directory.CreateDirectory(tempPath);
                foreach (var file in Directory.GetFiles(directory))
                    System.IO.File.Copy(file, Path.Combine(tempPath, Path.GetFileName(file)));
            }
        }
        public static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs=true,bool ignoreUpdater = false)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = new DirectoryInfo[]{};
            try
            {
                dirs = dir.GetDirectories();
            }
            catch( Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

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
                try 
                {
                    DateTime fstime = System.IO.File.GetLastWriteTime(file.Name);
                    DateTime fdtime = System.IO.File.GetLastWriteTime(temppath);
                    if (fstime!=fdtime) file.CopyTo(temppath, true); //copy if the files differ in time stamp
                }
                catch (Exception ex) { }
                //if (file.Name.Contains(".exe")) MessageBox.Show(sourceDirName + "       " + destDirName + "     " + file.Name);
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

            using (StreamWriter writer = new StreamWriter(deskDir + "\\" + Path.GetFileNameWithoutExtension(linkName) + ".url"))
            {
                string app = linkName;//System.Reflection.Assembly.GetExecutingAssembly().Location;
                writer.WriteLine("[InternetShortcut]");
                writer.WriteLine("URL=file:///" + app);
                writer.WriteLine("IconIndex=0");
                string icon = app.Replace('\\', '/');
                writer.WriteLine("IconFile=" + icon);
                writer.Flush();
            }
        }

        public static void CreateShortcut(string linkName)
        {
            object shDesktop = (object)"Desktop";
            WshShell shell = new WshShell();
            string shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop) + @"\" + Path.GetFileNameWithoutExtension(linkName) + ".lnk";
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
            shortcut.Description = "New shortcut for a " + Path.GetFileNameWithoutExtension(linkName);
            shortcut.Hotkey = "Ctrl+Alt+C";
            shortcut.TargetPath = linkName;
            shortcut.Save();
        }

    }
    public class Settings
    {
        public static string FolderStructure = @"SummaryTool\Software\";
        public string ITCDirectory = Path.Combine(@"\\flrblr001\windows_data\From-SLI-Fileserver\E\Product Engineering\CHAR DATA\", FolderStructure);
        public string AustinDirectory = Path.Combine(@"\\silabs.com\apps\Wireline\Timing\\", FolderStructure);
        //public string OtherReleasePath = new string() {"somePath"};
    }

}