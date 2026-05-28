namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.DocumentElements {
    public class PdfBookmarksModel : PdfModelBase {
        public PdfBookmarkOperationType OperationType { get; set; } = PdfBookmarkOperationType.Add;
        internal readonly List<FormatInfo> OperationTypes = new() {
            new() { Text = "Add", Value = PdfBookmarkOperationType.Add },
            new() { Text = "Delete", Value = PdfBookmarkOperationType.Delete }
        };
    }

    public enum PdfBookmarkOperationType {
        Add, Delete
    }
}
