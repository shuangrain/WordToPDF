using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordToPDF
{
    public partial class Form1 : Form
    {
        private readonly string[] _wordExt = new string[] { ".doc", ".docx" };

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
                lstFile.Items.AddRange(openFileDialog.FileNames.Where(x =>
                {
                    string ext = Path.GetExtension(x).ToLower();
                    return _wordExt.Any(y => y == ext);
                }).ToArray());
            }
        }

        private void btnStartConvert_Click(object sender, EventArgs e)
        {
            if (lstFile.Items.Count > 0)
            {
                //lock
                btnOpenFile.Enabled = false;
                btnStartConvert.Enabled = false;
                btnClear.Enabled = false;

                //run
                conventWordToPdf(lstFile.Items.Cast<string>().ToList()).ContinueWith((task) =>
                {
                    //unlock
                    MessageBox.Show("完成！");
                    btnOpenFile.Enabled = true;
                    btnStartConvert.Enabled = true;
                    btnClear.Enabled = true;
                }, TaskScheduler.FromCurrentSynchronizationContext());

            }
            else
            {
                MessageBox.Show("請選擇檔案!!", "Error");
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            lstFile.Items.Clear();
        }

        private Task conventWordToPdf(List<string> list)
        {
            return Task.Factory.StartNew(() =>
            {
                PDFConverter pdfConverter = new PDFConverter();
                list.AsParallel()
                    .ForAll(item =>
                    {
                        byte[] pdf = pdfConverter.GetPDF(item);
                        string ext = Path.GetExtension(item);
                        File.WriteAllBytes(item.Replace(ext, ".pdf"), pdf);

                        BeginInvoke(new Action(() =>
                        {
                            lstFile.Items.Remove(item);
                        }));
                    });

                pdfConverter.Dispose();
                pdfConverter = null;
            });

        }
    }
}
