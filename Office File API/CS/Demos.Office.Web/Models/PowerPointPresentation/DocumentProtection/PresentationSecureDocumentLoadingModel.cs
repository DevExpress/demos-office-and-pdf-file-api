using System.ComponentModel.DataAnnotations;
using DevExpress.Docs.Office;

namespace BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentProtection {
    public class PresentationSecureDocumentLoadingModel : PresentationModelBase {
        [RegularExpression(@"^(pptx|pptm|potx|potm)$",
            ErrorMessage = "Invalid output format. Supported output formats: pptx, pptm, potx, potm")]
        public override string OutputFormat { get; set; } = "pptx";
        public override string OutputName { get; set; } = "Secure_Presentation";

        // Loading limits
        [Range(1, 10, ErrorMessage = "Please enter a value between 1 and 10.")]
        public int MaxSlideCount { get; set; } = 7;
        [Range(1, 20, ErrorMessage = "Please enter a value between 1 and 20.")]
        public int MaxShapeCountPerSlide { get; set; } = 10;

        // Loading options
        public bool RemoveActiveXContent { get; set; } = true;
        public bool RemoveCustomXMLParts { get; set; } = true;
        public bool RemoveEmbeddedChartData { get; set; } = true;
        public bool RemoveExternalImages { get; set; } = true;
        public bool RemoveMacros { get; set; } = true;
        public bool RemoveOleObjects { get; set; } = true;
        public RestrictedHyperlinkRemovalMode RestrictedHyperlinkRemovalMode { get; set; } = RestrictedHyperlinkRemovalMode.UriOnly;

        public List<FormatInfo> OutputFormats { get; set; } = new() {
            new FormatInfo() { Text = "PowerPoint Presentation (*.pptx)", Value = "pptx" },
            new FormatInfo() { Text = "PowerPoint Macro-Enabled Presentation (*.pptm)", Value = "pptm" },
            new FormatInfo() { Text = "PowerPoint Template (*.potx)", Value = "potx" },
            new FormatInfo() { Text = "PowerPoint Macro-Enabled Template (*.potm)", Value = "potm" },
        };
    }
}
