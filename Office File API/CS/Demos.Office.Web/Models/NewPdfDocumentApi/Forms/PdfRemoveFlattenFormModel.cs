using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.Forms {
    public class PdfRemoveFlattenFormModel : PdfModelBase {
        public RemoveFlattenFormOperationType OperationType { get; set; }
        internal readonly List<FormatInfo> OperationTypes = new() {
            new() { Text = "Remove", Value = RemoveFlattenFormOperationType.Remove },
            new() { Text = "Flatten", Value = RemoveFlattenFormOperationType.Flatten }
        };
        public bool ProcessAll { get; set; } = true;
        public string FieldsToProcess { get; set; } = string.Empty;
    }

    public enum RemoveFlattenFormOperationType {
        Remove, Flatten
    }
}
