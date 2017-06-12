using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace WordToPDF
{
    /// <summary>
    /// http://blog.darkthread.net/post-2013-05-31-word-to-pdf.aspx
    /// </summary>
    public class PDFConverter
    {
        private Application wordApp = null;

        public PDFConverter()
        {
            wordApp = new Application();
            wordApp.Visible = false;
        }

        public byte[] GetPDF(string templateFile)
        {
            object filePath = templateFile;
            //檔案先寫入系統暫存目錄
            object outFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
            Document doc = null;
            try
            {
                object readOnly = true;
                doc = wordApp.Documents.Open(FileName: ref filePath, ReadOnly: ref readOnly);
                doc.Activate();
                //REF: http://bit.ly/Z9G5zg
                Range tmpRange = doc.Content;
                //去除醒目提示(Highlight)
                tmpRange.Find.Replacement.Highlight = 0;
                tmpRange.Find.Wrap = WdFindWrap.wdFindContinue;
                object replaceAll = WdReplace.wdReplaceAll;
                //釋放Range COM+                
                Marshal.FinalReleaseComObject(tmpRange);
                tmpRange = null;
                //存成PDF檔案
                object fileFormat = WdSaveFormat.wdFormatPDF;
                doc.SaveAs(FileName: ref outFile, FileFormat: ref fileFormat);
                //關閉Word檔
                object dontSave = WdSaveOptions.wdDoNotSaveChanges;
                doc.Close(ref dontSave);
            }
            finally
            {
                //確保Document COM+釋放
                if (doc != null)
                {
                    Marshal.FinalReleaseComObject(doc);
                }
                doc = null;
            }
            //讀取PDF檔，並將暫存檔刪除
            byte[] buff = File.ReadAllBytes(outFile.ToString());
            File.Delete(outFile.ToString());
            return buff;
        }

        public void Dispose()
        {
            //確實關閉Word Application
            try
            {
                object dontSave = WdSaveOptions.wdDoNotSaveChanges;
                wordApp.Quit(ref dontSave);
            }
            finally
            {
                Marshal.FinalReleaseComObject(wordApp);
                GC.Collect();
            }
        }
    }
}
