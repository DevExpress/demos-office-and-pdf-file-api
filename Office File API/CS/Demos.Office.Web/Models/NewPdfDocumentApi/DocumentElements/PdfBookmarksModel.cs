using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentElements {
    public class PdfBookmarksModel : PdfModelBase {
        public NewPdfBookmarkOperationType OperationType { get; set; } = NewPdfBookmarkOperationType.Add;
        internal readonly List<FormatInfo> OperationTypes = new() {
            new() { Text = "Add", Value = NewPdfBookmarkOperationType.Add },
            new() { Text = "Delete", Value = NewPdfBookmarkOperationType.Delete }
        };
    }

    public enum NewPdfBookmarkOperationType {
        Add, Delete
    }
}
