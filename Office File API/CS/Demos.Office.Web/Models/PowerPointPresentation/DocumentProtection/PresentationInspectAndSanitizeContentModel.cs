namespace BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentProtection {
    public class PresentationInspectAndSanitizeContentModel : PresentationModelBase {
        public string OutputFileName {
            get { return OperationType == InspectOperationType.Inspect ? "Inspection_Log" : "Sanitized_Presentation"; }
        }
        public string OutputFileFormat {
            get { return OperationType == InspectOperationType.Inspect ? "txt" : "pptx"; }
        }

        // Hidden content
        public bool RemoveHiddenSlides { get; set; } = true;
        public bool RemoveHiddenShapes { get; set; } = true;

        // Sensitive content
        public bool RemoveActiveXContent { get; set; } = true;
        public bool RemoveComments { get; set; } = true;
        public bool RemoveCustomXMLParts { get; set; } = true;
        public bool RemoveEmbeddedChartData { get; set; } = true;
        public bool RemoveNotes { get; set; } = true;
        public bool RemoveMacros { get; set; } = true;
        public bool RemoveMetadata { get; set; } = true;
        public bool RemoveOleObjects { get; set; } = true;
        public bool RemoveOffSlideContent { get; set; } = true;

        public InspectOperationType OperationType { get; set; } = InspectOperationType.Inspect;

        public enum InspectOperationType {
            Inspect,
            Sanitize
        }
    }
}
