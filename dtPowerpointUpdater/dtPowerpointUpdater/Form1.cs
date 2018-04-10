using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.PowerPoint;

namespace dtPowerpointUpdater
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string pathToWatch = "";
        Microsoft.Office.Interop.PowerPoint.Application pptApp;
        Microsoft.Office.Interop.PowerPoint.Presentations ps;
        Microsoft.Office.Interop.PowerPoint.Presentation p;

        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (var process in Process.GetProcessesByName("POWERPNT"))
            {
                process.Kill();
            }
            Thread.Sleep(1000);
            this.WindowState = FormWindowState.Minimized;
            timer1.Interval = 1000;
            timer1.Start();
            pathToWatch = fileLocationTxt.Text;
            //watch();

            try
            {
               
                showPowerpoint();

            }
            catch { }

        }
        private void watch()
        {
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = pathToWatch;
      

            watcher.Filter = "*.pptx*";
            watcher.Changed += new FileSystemEventHandler(OnChanged);
            watcher.EnableRaisingEvents = true;
        }

        private void OnChanged(object source, FileSystemEventArgs e)
        {

        
    

        }

        private void updateFileLocationBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {                
                try
                {
                    string file = folderBrowserDialog1.SelectedPath;
                    string fullPath = Path.GetFullPath(file);
                    fileLocationTxt.Text = fullPath;
                    pathToWatch = fullPath;
                }
                catch
                {
                    MessageBox.Show("Error! Can't do that.");
                }
                showPowerpoint();
            }
           
         
        }

       
       
        void showPowerpoint()
        {

        
            pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
            var directory = new DirectoryInfo(fileLocationTxt.Text);
            var myFile = directory.GetFiles()
             .OrderByDescending(f => f.LastWriteTime)
             .First();
            string path = Path.Combine(myFile.Directory.FullName,myFile.Name);

            pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            pptApp.Activate();

           ps = pptApp.Presentations;
            try
            {
                p = ps.Open(path,
                           Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue);
       
                pptApp.ActivePresentation.SlideShowSettings.Run();
            }
            catch { }
        }

        private void fileLocationTxt_TextChanged(object sender, EventArgs e)
        {

        }
        int oldcount = 0;
        bool isRunning = false;
        private void timer1_Tick(object sender, EventArgs e)
        {
       
            int count = Directory.GetFiles(pathToWatch).Length;

            if (count != oldcount)
            {
                oldcount = count;
                    pptApp.ActivePresentation.Close();
                    // pptApp.Quit();
                    showPowerpoint();               
            }
            else
            {
                if (isRunning == false)
                {

                }
                else
                {

                }

            }


        }
    }
}
