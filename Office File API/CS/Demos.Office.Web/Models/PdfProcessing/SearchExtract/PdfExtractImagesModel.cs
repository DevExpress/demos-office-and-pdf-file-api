using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.SearchExtract {
    public class PdfExtractImagesModel : PdfPageRangeModelBase {
        internal readonly List<FormatInfo> ImageFormats = new() {
            new() { Text = "PNG (Portable Network Graphics)", Value = "png" },
            new() { Text = "JPEG (Joint Photographic Experts Group)", Value = "jpeg" },
            new() { Text = "BMP (Bitmap)", Value = "bmp" },
            new() { Text = "TIFF (Tagged Image File Format)", Value = "tiff" },
            new() { Text = "SVG (Scalable Vector Graphics)", Value = "svg" }
        };
        public string ImageFormat { get; set; } = "png";
        public override string OutputFormat { get => "zip"; }
        public override string OutputName { get => "Images"; }
        [Range(72, 1200, ErrorMessage = "Resolution must be between 72 and 1200 DPI")]
        public int Resolution { get; set; } = 300;
        internal readonly List<int> PredefinedResolutions = new() { 72, 96, 300 };
    }
}
