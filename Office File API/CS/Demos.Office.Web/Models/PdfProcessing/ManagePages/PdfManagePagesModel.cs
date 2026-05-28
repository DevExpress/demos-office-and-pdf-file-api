using DevExpress.Pdf;

namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.ManagePages {
    public class PdfManagePagesModel : PdfPageRangeModelBase {
        public PdfManagePagesModel() {
            ExportMode = RangeExportMode.First;
        }
        internal override List<object> ExportModes { get; } = new() {
            new { Text = "First Page Only", Value = RangeExportMode.First },
            new { Text = "Custom Page Range", Value = RangeExportMode.CustomRange }
        };
        internal List<FormatInfo> OperationsTypes { get; } = new() {
            new() { Text = "Add a New Page", Value = PdfPageOperationType.Add },
            new() { Text = "Delete a Page", Value = PdfPageOperationType.Delete },
            new() { Text = "Duplicate a Page", Value = PdfPageOperationType.Duplicate }
        };
        public PdfRectangle NewPageSize { get; set; } = PdfPaperSize.A4;
        internal List<object> NewPageSizes { get; } = new() {
            new { Text = "A2", Value = PdfPaperSize.A2 },
            new { Text = "A3", Value = PdfPaperSize.A3 },
            new { Text = "A4", Value = PdfPaperSize.A4 },
            new { Text = "A5", Value = PdfPaperSize.A5 },
            new { Text = "Letter", Value = PdfPaperSize.Letter }
        };
        public PdfPageOperationType OperationType { get; set; } = PdfPageOperationType.Add;
        public double Width { get; set; } = 0;
        public double Height { get; set; } = 0;
    }

    public enum PdfPageOperationType {
        Add, Delete, Duplicate
    }
}
