using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordToPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word 檔(*.docx,*.doc)|*.docx;*.doc";
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                for (int i = 0; i < openFileDialog.FileNames.Length; i++)
                {
                    String filePath = openFileDialog.FileNames.GetValue(i).ToString();
                    lstFile.Items.Add(filePath);
                }
            }
        }

        private void btnStartConvert_Click(object sender, EventArgs e)
        {
            if (lstFile.Items.Count > 0)
            {
                backgroundWorker1.RunWorkerAsync();
                btnOpenFile.Enabled = false;
                btnStartConvert.Enabled = false;
                btnClear.Enabled = false;
            }
            else
            {
                MessageBox.Show("請選擇檔案!!", "Error");
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            PdfConverter pdfConverter = new PdfConverter();
            for (int i = 0; i < lstFile.Items.Count; i++)
            {
                String tmp_FilePath = lstFile.Items[i].ToString();
                var pdf = pdfConverter.GetPDF(tmp_FilePath);
                if (tmp_FilePath.Split('.')[1] == "doc")
                {
                    tmp_FilePath.Replace(".doc", "");

                    tmp_FilePath.Substring(0, tmp_FilePath.Length - 4);
                }
                else if (tmp_FilePath.Split('.')[1] == "docx")
                {
                    tmp_FilePath.Replace(".docx", "");
                }
                System.IO.File.WriteAllBytes(tmp_FilePath + ".pdf", pdf);
            }
            pdfConverter.Dispose();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            lstFile.Items.Clear();
            MessageBox.Show("完成！");
            btnOpenFile.Enabled = true;
            btnStartConvert.Enabled = true;
            btnClear.Enabled = true;
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            lstFile.Items.Clear();
        }
    }
}
