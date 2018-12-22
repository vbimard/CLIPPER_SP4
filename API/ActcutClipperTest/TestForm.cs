using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

using Actcut.ActcutClipperApi;

namespace ActcutClipperTest
{
    public partial class TestForm : Form
    {
        private ClipperApi ClipperApi = null;
        private Process QuoteModelApiProcess = null;

        private string ParamFolder = null;
        private string ParamFile = null;
        private string ResultFile = null;

        private string AlmaCamDB = null;
        private string User = null;
        private long quoteNumber = 0;

        public TestForm()
        {
            InitializeComponent();

            ParamFolder = Program.AlmaCamBinFolder + @"ActcutClipperExeParam"; // ActcutClipperExeParam : nom de sous répertoire imposé
            ParamFile = ParamFolder + @"\Param.txt"; // Nom du fichier contenant les commandes : nom imposé
            ResultFile = ParamFolder + @"\Result.txt"; // Nom du fichier contenant les resultats : nom imposé

            AlmaCamDB = "AlmaCam_Clipper_8_Sp4";
            User = "SUPER";
        }

        #region Api Test

        private void BtnInitApi_Click(object sender, EventArgs e)
        {
            ClipperApi = new ClipperApi();
            ClipperApi.Init(AlmaCamDB, User);
        }
        private void BtnGetQuoteApi_Click(object sender, EventArgs e)
        {
            long quoteNumberReference = -1;
            ClipperApi.GetQuote(out quoteNumberReference);
            quoteNumber = quoteNumberReference;
        }
        private void BtnExportQuoteApi_Click(object sender, EventArgs e)
        {
            ClipperApi.ExportQuote(quoteNumber, "ALMA", @"C:\Temp\"+quoteNumber+".txt");
        }

        #endregion

        #region Exe Test

        private void BtnInit_Click(object sender, EventArgs e)
        {
            Process[] processList = Process.GetProcesses();
            foreach (Process process in processList)
            {
                //if (process.ProcessName == "Actcut.ActcutClipperExe")
                if (process.ProcessName == "ActcutClipperExe")
                {
                    QuoteModelApiProcess = process;
                    MessageBox.Show("Actcut.ActcutClipperExe deja lancé");
                    return;
                }
            }

            if (File.Exists(ResultFile)) File.Delete(ResultFile);

            QuoteModelApiProcess = new System.Diagnostics.Process();
            //QuoteModelApiProcess.StartInfo.FileName = Path.Combine(Program.AlmaCamBinFolder, "Actcut.ActcutClipperExe.exe");
            //QuoteModelApiProcess.StartInfo.FileName = Path.Combine(Program.AlmaCamBinFolder, "Actcut.ActcutClipperExe.exe");
            QuoteModelApiProcess.StartInfo.FileName = Path.Combine(Program.AlmaCamBinFolder, "AlmaCamUser1.exe");
            QuoteModelApiProcess.StartInfo.Arguments = "Action=Init User=" + User + " Db=" + AlmaCamDB;
            QuoteModelApiProcess.Start();

            WaitResult(ResultFile);
        }
        private void BtnGetQuote_Click(object sender, EventArgs e)
        {
            if (File.Exists(ParamFile)) File.Delete(ParamFile);
            if (File.Exists(ResultFile)) File.Delete(ResultFile);

            File.WriteAllText(ParamFile, @"Action=GetQuote");
            WaitResult(ResultFile);
        }
        private void BtnExportQuote_Click(object sender, EventArgs e)
        {
            if (File.Exists(ParamFile)) File.Delete(ParamFile);
            if (File.Exists(ResultFile)) File.Delete(ResultFile);

            File.WriteAllText(ParamFile, @"Action=ExportQuote QuoteNumber=7 OrderNumber=987 ExportFile=C:\Temp\toto.txt");
            WaitResult(ResultFile);
        }
        private void BtnExit_Click(object sender, EventArgs e)
        {
            if (File.Exists(ParamFile)) File.Delete(ParamFile);

            File.WriteAllText(ParamFile, @"Action=Exit");
        }

        private void WaitResult(string resultFile)
        {
            while (File.Exists(resultFile) == false && QuoteModelApiProcess != null && QuoteModelApiProcess.HasExited == false)
                System.Threading.Thread.Sleep(10);

            if (File.Exists(resultFile))
            {
                string result = File.ReadAllText(resultFile);
                MessageBox.Show(result);
            }
            else
            {
                MessageBox.Show("Error");
            }
        }

        #endregion
    }
}
