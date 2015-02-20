using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace VendorEDI
{
    public partial class FormMain : Form
    {
        private List<string> ediFiles;

        public FormMain()
        {
            InitializeComponent();

            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(FormMain_DragEnter);
            this.DragDrop += new DragEventHandler(FormMain_DragDrop);
        }

        private void btnDone_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        void FormMain_DragDrop(object sender, DragEventArgs e)
        {
            // Don't allow new files if we are already working on some.
            if (bw.IsBusy)
            {
                return;
            }

            ediFiles = new List<string>();
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                ediFiles.Add(file);
            }

            if (!bw.IsBusy)
            {
                bw.RunWorkerAsync();
            }
        }

        void FormMain_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }
        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = bw.IsBusy;
        }

        #region Invoke Required

        void EnableBtnDone(bool enabled)
        {
            if (this.btnDone.InvokeRequired)
            {
                this.BeginInvoke(new MethodInvoker(() => EnableBtnDone(enabled)));
            }
            else
            {
                this.btnDone.Enabled = enabled;
            }
        }

        #endregion

        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            EnableBtnDone(false);

            var db = new DbUtilities();
            var worker = sender as BackgroundWorker;

            foreach (var fileName in ediFiles)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                db.Initialize(fileName);
                // string csvName = FileUtilities.CleanFile(fileName);
            }
        }

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
        }

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            EnableBtnDone(true);
        }
    }
}
