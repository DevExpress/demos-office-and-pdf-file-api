namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.SearchExtract {
    public class PdfXMPMetadataModel : PdfModelBase {
        internal List<FormatInfo> MetadataOperations { get; set; } = new List<FormatInfo>() {
            new() { Text = "Extract", Value = MetadataOperation.Extract },
            new() { Text = "Add", Value = MetadataOperation.Add },
            new() { Text = "Remove", Value = MetadataOperation.Remove }
        };
        public MetadataOperation MetadataOperation { get; set; } = MetadataOperation.Extract;
        internal List<string> AllowedMetadataExtensions { get; set; } = new() {
            ".xmp"
        };
    }

    public enum MetadataOperation {
        Extract, Add, Remove
    }
}
