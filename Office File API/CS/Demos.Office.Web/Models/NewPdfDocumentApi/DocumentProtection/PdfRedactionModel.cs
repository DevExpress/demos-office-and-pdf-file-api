using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentProtection {
    public class PdfRedactionModel : PdfModelBase {
        public bool ApplyRedaction { get; set; } = true;
        [Required(ErrorMessage = "Add a search term for sensitive data")]
        [MaxLength(15, ErrorMessage = "Text must not exceed 15 symbols")]
        public string SearchText { get; set; } = "DevExpress";
        [MaxLength(15, ErrorMessage = "Text must not exceed 15 symbols")]
        public string RedactionText { get; set; } = "REDACTED";
    }
}
