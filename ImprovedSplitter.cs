using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iText.Kernel.Pdf;
using iText.Kernel.Utils;

namespace CieIntercalaImpresion
{
    class ImprovedSplitter : PdfSplitter
    {
        private Func<PageRange, PdfWriter> nextWriter;
        public ImprovedSplitter(PdfDocument pdfDocument, Func<PageRange, PdfWriter> nextWriter) : base(pdfDocument)
        {
            this.nextWriter = nextWriter;
        }

        protected override PdfWriter GetNextPdfWriter(PageRange documentPageRange)
        {
            return nextWriter.Invoke(documentPageRange);
        }
    }
}
