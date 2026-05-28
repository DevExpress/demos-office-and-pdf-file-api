using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.ManagePages {
    public class PdfManagePagesModel : PdfPageRangeModelBase {
        public PdfManagePagesModel() {
            ExportMode = RangeExportMode.First;
        }
        internal override List<object> ExportModes { get; } = new() {
            new { Text = "First Page Only",    Value = RangeExportMode.First },
            new { Text = "Custom Page Range",  Value = RangeExportMode.CustomRange }
        };
        internal List<FormatInfo> OperationsTypes { get; } = new() {
            new() { Text = "Add a New Page",   Value = NewPdfPageOperationType.Add },
            new() { Text = "Delete a Page",    Value = NewPdfPageOperationType.Delete },
            new() { Text = "Duplicate a Page", Value = NewPdfPageOperationType.Duplicate }
        };
        public NewPageSize NewPageSize { get; set; } = NewPageSize.A4;
        internal List<object> NewPageSizes { get; } = new() {
            new { Text = "A2", Value = NewPageSize.A2 },
            new { Text = "A3", Value = NewPageSize.A3 },
            new { Text = "A4", Value = NewPageSize.A4 },
            new { Text = "A5", Value = NewPageSize.A5 },
            new { Text = "Letter", Value = NewPageSize.Letter }
        };
        public NewPdfPageOperationType OperationType { get; set; } = NewPdfPageOperationType.Add;
        public double Width  { get; set; } = 0;
        public double Height { get; set; } = 0;
    }

    public enum NewPdfPageOperationType { Add, Delete, Duplicate }

    public enum NewPageSize {
        A2, A3, A4, A5, Letter
    }

    public static class NewPageSizeHelper {
        public static (double Width, double Height) GetDimensions(NewPageSize size) => size switch {
            NewPageSize.A2 => (1190.55, 1683.78),
            NewPageSize.A3 => (841.89, 1190.55),
            NewPageSize.A4 => (595.28, 841.89),
            NewPageSize.A5 => (419.53, 595.28),
            _              => (595.28, 841.89)
        };
    }
}
