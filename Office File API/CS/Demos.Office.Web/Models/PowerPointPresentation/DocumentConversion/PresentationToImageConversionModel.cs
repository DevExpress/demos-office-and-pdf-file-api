using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;
using DevExpress.Drawing;

namespace BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentConversion {
    public class PresentationToImageConversionModel : PresentationModelBase, IValidatableObject {
        public DXImageFormat ImageFormat { get; set; } = DXImageFormat.Png;

        [Range(72, 1200, ErrorMessage = "Resolution must be between 72 and 1200 DPI")]
        public int Resolution { get; set; } = 300;
        public RangeExportMode ExportMode { get; set; } = RangeExportMode.All;

        public string SlideRange { get; set; } = string.Empty;

        public override string OutputName { get; set; } = "Document_Images";
        internal readonly List<FormatInfo> ImageFormats = new() {
            new FormatInfo { Text = "PNG (Portable Network Graphics)", Value = DXImageFormat.Png },
            new FormatInfo { Text = "JPEG (Joint Photographic Experts Group)", Value = DXImageFormat.Jpeg },
            new FormatInfo { Text = "BMP (Bitmap)", Value = DXImageFormat.Bmp },
            new FormatInfo { Text = "SVG (Scalable Vector Graphics)", Value = DXImageFormat.Svg }
        };

        internal readonly List<object> ExportModes = new() {
            new { Text = "First Slide Only", Value = RangeExportMode.First },
            new { Text = "All Slides", Value = RangeExportMode.All },
            new { Text = "Custom Slide Range", Value = RangeExportMode.CustomRange }
        };

        public bool IsSlideRangeEnabled => ExportMode == RangeExportMode.CustomRange;

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            if(IsSlideRangeEnabled) {
                if(string.IsNullOrWhiteSpace(SlideRange)) {
                    yield return new ValidationResult(
                        "SlideRange is required when Export Mode is 'custom'.",
                        new[] { nameof(SlideRange) });
                    yield break;
                }

                // Accept patterns like: 1,2,4-6, 10-12,15
                const string pattern = @"^\s*\d+(\s*-\s*\d+)?(\s*,\s*\d+(\s*-\s*\d+)?)*\s*$";
                if(!Regex.IsMatch(SlideRange, pattern)) {
                    yield return new ValidationResult(
                        "SlideRange must be a comma-separated list of numbers or ranges (e.g. 1,3,5-7, 10-12).",
                        new[] { nameof(SlideRange) });
                }
            }
        }
    }
}
