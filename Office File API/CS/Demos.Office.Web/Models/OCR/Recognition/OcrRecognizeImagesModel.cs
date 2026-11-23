namespace BlazorDemo.OfficeFileApi.Models.OCR.Recognition {
    public class OcrRecognizeImagesModel : FileModelBase {
        public override string OutputFormat { get; set; } = "pdf";
        public override string OutputName { get; set; } = "RecognizedImage";
        public override List<string> AllowedFileExtensions { get; set; } = new() { ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff" };
        public bool DetectPageOrientation { get; set; } = true;

        internal readonly List<FormatInfo> OutputFormats = new() {
            new() { Text = "PDF (Searchable Document)", Value = "pdf" },
            new() { Text = "TXT (Plain Text)", Value = "txt" },
            new() { Text = "JSON (Text Fragments)", Value = "json" }
        };
    }
}
