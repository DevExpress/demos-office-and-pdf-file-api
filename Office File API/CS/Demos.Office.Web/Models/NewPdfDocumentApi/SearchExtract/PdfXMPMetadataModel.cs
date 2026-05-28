using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.SearchExtract {
    public class PdfXMPMetadataModel : PdfModelBase {
        internal List<FormatInfo> MetadataOperations { get; set; } = new() {
            new() { Text = "Extract", Value = NewMetadataOperation.Extract },
            new() { Text = "Add",     Value = NewMetadataOperation.Add },
            new() { Text = "Remove",  Value = NewMetadataOperation.Remove }
        };
        public NewMetadataOperation MetadataOperation { get; set; } = NewMetadataOperation.Extract;
    }

    public enum NewMetadataOperation {
        Extract, Add, Remove
    }
}
