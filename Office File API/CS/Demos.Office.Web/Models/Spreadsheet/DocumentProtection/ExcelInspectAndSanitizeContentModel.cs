namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentProtection {
    public class ExcelInspectAndSanitizeContentModel : ExcelFileModelBase {
        public string OutputFileName {
            get { return OperationType == InspectOperationType.Inspect ? "Inspection_Log" : "Sanitized_Document"; }
        }
        public string OutputFileFormat {
            get { return OperationType == InspectOperationType.Inspect ? "txt" : "xlsx"; }
        }

        // Hidden content
        public bool RemoveHiddenRows { get; set; } = true;
        public bool RemoveHiddenColumns { get; set; } = true;
        public bool RemoveHiddenSheets { get; set; } = true;
        public bool RemoveInvisibleCellText { get; set; } = true;

        // Sensitive content
        public bool RemoveMetadata { get; set; } = true;
        public bool RemoveComments { get; set; } = true;
        public bool RemoveCustomXMLParts { get; set; } = true;
        public bool RemoveActiveXContent { get; set; } = true;
        public bool RemoveMacros { get; set; } = true;
        public bool RemoveOleObjects { get; set; } = true;
        public bool RemoveThreadedComments { get; set; } = true;

        public InspectOperationType OperationType { get; set; } = InspectOperationType.Inspect;

        public enum InspectOperationType {
            Inspect,
            Sanitize
        }
    }
}
