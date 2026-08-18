using iText.Kernel.Events;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;

namespace PACKINGSEAL.Utils
{
    public class HeaderFooterHandler : IEventHandler
    {
        private readonly int totalPages;
        private readonly string today;

        public HeaderFooterHandler(int totalPages, string today)
        {
            this.totalPages = totalPages;
            this.today = today;
        }

        public void HandleEvent(Event @event)
        {
            var docEvent = (PdfDocumentEvent)@event;
            var pdfDoc = docEvent.GetDocument();
            var page = docEvent.GetPage();
            int pageNumber = pdfDoc.GetPageNumber(page);

            PdfCanvas pdfCanvas = new PdfCanvas(page.NewContentStreamAfter(), page.GetResources(), pdfDoc);
            Canvas canvas = new Canvas(pdfCanvas, page.GetPageSize());

            // Vẽ header góc phải
            canvas.ShowTextAligned(
                 new Paragraph($"Page {pageNumber}/{totalPages}  Date: {today}")
                .SetFontSize(10),
                page.GetPageSize().GetWidth() - 40,
                page.GetPageSize().GetTop() - 30,
                TextAlignment.RIGHT
            );

            canvas.Close();
        }
    }

}
