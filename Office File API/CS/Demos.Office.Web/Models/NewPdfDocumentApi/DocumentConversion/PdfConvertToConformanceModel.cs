using BlazorDemo.OfficeFileApi.Models.PdfProcessing;
using DevExpress.Docs.Pdf;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentConversion {
    public class PdfConvertToConformanceModel : PdfModelBase {
        internal readonly List<FormatInfo> CompatibilityLevels = new() {
            new() { Text = "PDF/A-1b", Value = PdfACompatibility.PdfA1b },
            new() { Text = "PDF/A-2b", Value = PdfACompatibility.PdfA2b },
            new() { Text = "PDF/A-3b", Value = PdfACompatibility.PdfA3b }
        };
        public PdfACompatibility Compatibility { get; set; } = PdfACompatibility.PdfA2b;
    }
}
