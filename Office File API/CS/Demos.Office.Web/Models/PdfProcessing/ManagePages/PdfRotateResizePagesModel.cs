using DevExpress.Pdf;

namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.ManagePages {
    public class PdfRotateResizePagesModel : PdfPageRangeModelBase {
        public int Rotate { get; set; } = 0;
        public PdfRectangle PageSize { get; set; } = PdfPaperSize.A4;
        public double Width { get; set; } = 0;
        public double Height { get; set; } = 0;
        internal List<object> PageSizes { get; } = new() {
            new { Text = "A2", Value = PdfPaperSize.A2 },
            new { Text = "A3", Value = PdfPaperSize.A3 },
            new { Text = "A4", Value = PdfPaperSize.A4 },
            new { Text = "A5", Value = PdfPaperSize.A5 },
            new { Text = "Letter", Value = PdfPaperSize.Letter }
        };
        internal List<object> RotationAngles { get; } = new() {
            new { Text = "0°", Value = 0 },
            new { Text = "90°", Value = 90 },
            new { Text = "180°", Value = 180 },
            new { Text = "270°", Value = 270 }
        };
    }
}
