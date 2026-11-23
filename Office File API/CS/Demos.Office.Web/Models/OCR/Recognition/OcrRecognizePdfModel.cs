using BlazorDemo.OfficeFileApi.Models.PdfProcessing;
using DevExpress.Docs.Ocr;

namespace BlazorDemo.OfficeFileApi.Models.OCR.Recognition {
    public class OcrRecognizePdfModel : PdfPageRangeModelBase {
        public override string OutputFormat { get; set; } = "pdf";
        public override string OutputName { get; set; } = "RecognizedPdf";
        public bool DetectPageOrientation { get; set; } = true;
        public PdfDocumentRecognitionMode RecognitionMode { get; set; } = PdfDocumentRecognitionMode.RenderedPage;

        internal readonly List<FormatInfo> OutputFormats = new() {
            new() { Text = "PDF (Searchable Document)", Value = "pdf" },
            new() { Text = "TXT (Plain Text)", Value = "txt" },
            new() { Text = "JSON (Text Fragments)", Value = "json" }
        };

        internal readonly List<object> RecognitionModes = new() {
            new { Text = "Rendered pages (any PDF)", Value = PdfDocumentRecognitionMode.RenderedPage },
            new { Text = "Embedded images (scanned PDF)", Value = PdfDocumentRecognitionMode.EmbeddedImages }
        };
    }
}
