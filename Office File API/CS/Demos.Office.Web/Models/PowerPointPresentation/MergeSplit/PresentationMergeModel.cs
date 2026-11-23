using DevExpress.Docs.Presentation;

namespace BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.MergeSplit {
    public class PresentationMergeModel : PresentationModelBase {
        public override string OutputName { get; set; } = "Merged_Presentation";
        public override string OutputFormat { get; set; } = "pdf";
        public bool IsPdfOutput { get { return OutputFormat?.ToLowerInvariant() == "pdf"; } }
        public ResizeMode ContentScaleMode { get; set; } = ResizeMode.EnsureFit;
        internal readonly List<object> ContentScaleModes = new() {
            new { Text = "Center", Value = ResizeMode.Center },
            new { Text = "Stretch", Value = ResizeMode.Stretch },
            new { Text = "Fit Entire Slide", Value = ResizeMode.EnsureFit },
            new { Text = "Fill Entire Slide", Value = ResizeMode.EnsureFill }
        };
        public List<FormatInfo> OutputFormats { get; set; } = new() {
            new FormatInfo { Text = "PDF Document (.pdf)", Value = "pdf" },
            new FormatInfo { Text = "PowerPoint Presentation (.pptx)", Value = "pptx" },
            new FormatInfo { Text = "PowerPoint Macro-Enabled Presentation (.pptm)", Value = "pptm" },
            new FormatInfo { Text = "PowerPoint Template (.potx)", Value = "potx" },
            new FormatInfo { Text = "PowerPoint Macro-Enabled Template (.potm)", Value = "potm" }
        };
    }
}
