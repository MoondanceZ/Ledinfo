using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.ComponentModel;
using FI.Public;

namespace KY.Fi.DCZqLQ
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Process[] curProcesses = Process.GetProcessesByName(Application.ProductName);
            if (curProcesses.Length > 1)
            {
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            System.Threading.Thread.Sleep(5000);
            Const.StartupPath = Path.GetDirectoryName(Application.ExecutablePath);
            Const.DBConnFile = Path.GetDirectoryName(Application.ExecutablePath) + "\\Resources\\DataFile.xml";
            CBase.GetSysParam(Const.DBConnFile);

            StartupManager startupManager = new StartupManager();
            if (Const.IsRunAsStart == "1")
            {
                startupManager.Startup = true;
            }
            else
            {
                startupManager.Startup = false;
            }
            if (Const.model == "自动运行")
            {
                BackgroundWorker bgThread = new BackgroundWorker();
                bgThread.WorkerSupportsCancellation = true;
                bgThread.DoWork += new DoWorkEventHandler(new AutoExportExcel().ExportExcel);
                bgThread.RunWorkerAsync();
            }
            using (DCZqIOMain form = new DCZqIOMain())
            {
                Application.Run();
            }

        }
    }
}
