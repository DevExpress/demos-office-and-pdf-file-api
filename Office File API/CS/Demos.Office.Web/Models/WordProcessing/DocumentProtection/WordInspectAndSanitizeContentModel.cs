namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentProtection {
    public class WordInspectAndSanitizeContentModel : WordFileModelBase {
        public string OutputFileName {
            get { return OperationType == InspectOperationType.Inspect ? "Inspection_Log" : "Sanitized_Document"; }
        }
        public string OutputFileFormat {
            get { return OperationType == InspectOperationType.Inspect ? "txt" : "docx"; }
        }

        // Hidden content
        public bool RemoveHiddenText { get; set; } = true;
        public bool RemoveInvisibleText { get; set; } = true;

        // Sensitive content
        public bool RemoveMetadata { get; set; } = true;
        public bool RemoveComments { get; set; } = true;
        public bool RemoveCustomXMLParts { get; set; } = true;
        public bool RemoveActiveXContent { get; set; } = true;
        public bool RemoveMacros { get; set; } = true;
        public bool RemoveOleObjects { get; set; } = true;
        public bool RemoveTrackChanges { get; set; } = true;

        public InspectOperationType OperationType { get; set; } = InspectOperationType.Inspect;

        public enum InspectOperationType {
            Inspect,
            Sanitize
        }
    }
}
