using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.OCR.Recognition {
    public enum OcrPageRegion {
        Header,
        Body,
        Footer
    }

    public class OcrRecognizeRegionModel : PdfModelBase {
        public override string OutputFormat { get; set; } = "pdf";
        public override string OutputName { get; set; } = "RecognizedRegion";
        public OcrPageRegion Region { get; set; } = OcrPageRegion.Header;

        internal readonly List<FormatInfo> OutputFormats = new() {
            new() { Text = "PDF (Searchable Document)", Value = "pdf" },
            new() { Text = "TXT (Plain Text)", Value = "txt" },
            new() { Text = "JSON (Text Fragments)", Value = "json" }
        };

        internal readonly List<object> Regions = new() {
            new { Text = "Header", Value = OcrPageRegion.Header },
            new { Text = "Body", Value = OcrPageRegion.Body },
            new { Text = "Footer", Value = OcrPageRegion.Footer }
        };
    }
}
