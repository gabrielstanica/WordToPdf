using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPdf
{
    class ActionsLogic
    {
        public Document WordDocument { get; set; }

        public void GetWord(string pathToDoc)
        {
            if(Path.GetExtension(pathToDoc).Contains(() => ))
            {
                // https://www.google.ro/search?q=convert+word+to+pdf+c%23&ie=utf-8&oe=utf-8&gws_rd=cr&ei=xHUPV92oOIryUJ3Xq5gD
                // https://code.msdn.microsoft.com/office/Word-file-to-PDF-Conversion-261fd865
                // https://msdn.microsoft.com/en-us/library/system.io.path.getextension%28v=vs.110%29.aspx
            }
        }

        public void Convert(string path)
        {
            Application wordApp = new Application();
            var wordDoc = wordApp.Documents.Open(path);

            wordDoc.ExportAsFixedFormat(path, WdExportFormat.wdExportFormatPDF);
        }

    }
}
