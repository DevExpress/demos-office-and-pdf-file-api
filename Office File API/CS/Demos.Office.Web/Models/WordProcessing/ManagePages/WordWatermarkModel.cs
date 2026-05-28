using DevExpress.XtraRichEdit.API.Native;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.ManagePages
{
    public class WordWatermarkModel : WordFileModelBase, IValidatableObject
    {
        // Watermark type
        [Required(ErrorMessage = "Watermark type is required")]
        public WatermarkType WatermarkType { get; set; } = WatermarkType.Text;

        // Text watermark properties
        [Required(ErrorMessage = "Watermark text is required")]
        public string WatermarkText { get; set; } = "CONFIDENTIAL";

        //[Required]
        public string TextColor { get; set; }
        internal object Color { get; set; } = "#808080"; // Default gray color

        [Required(ErrorMessage = "Font name is required")]
        public string FontName { get; set; } = "Calibri";

        [Range(1, 144, ErrorMessage = "Font size must be between 1 and 144 points")]
        public int FontSize { get; set; } = 72;

        [Required(ErrorMessage = "Layout is required")]
        public WatermarkLayout Layout { get; set; } = WatermarkLayout.Diagonal;

        public bool Semitransparent { get; set; } = true;

        [Range(0.0, 65.5, ErrorMessage = "Scale factor must be between 0 (0%) and 65.5 (6550%)")]
        public double ScaleFactor { get; set; } = 1.0; // 1.0 = 100%

        public bool Washout { get; set; } = true;

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "Watermarked_Document";

        internal readonly List<object> WatermarkTypes = new()
        {
            new { Text = "Text Watermark", Value = WatermarkType.Text },
            new { Text = "Image Watermark", Value = WatermarkType.Image }
        };

        internal readonly List<object> FontNames = new()
        {
            new { Text = "Calibri", Value = "Calibri" },
            new { Text = "Arial", Value = "Arial" },
            new { Text = "Times New Roman", Value = "Times New Roman" },
            new { Text = "Verdana", Value = "Verdana" },
            new { Text = "Tahoma", Value = "Tahoma" },
            new { Text = "Georgia", Value = "Georgia" },
            new { Text = "Impact", Value = "Impact" }
        };

        internal readonly List<object> Layouts = new()
        {
            new { Text = "Diagonal", Value = WatermarkLayout.Diagonal },
            new { Text = "Horizontal", Value = WatermarkLayout.Horizontal }
        };

        internal readonly List<FormatInfo> OutputFormats = new()
        {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            var results = new List<ValidationResult>();

            // Validate watermark type
            var validTypes = new[] { WatermarkType.Text, WatermarkType.Image };
            if (!validTypes.Contains(WatermarkType))
            {
                results.Add(new ValidationResult(
                    "Invalid watermark type",
                    new[] { nameof(WatermarkType) }));
            }

            // Validate text watermark settings
            if (WatermarkType == WatermarkType.Text)
            {
                if (string.IsNullOrWhiteSpace(WatermarkText))
                {
                    results.Add(new ValidationResult(
                        "Watermark text is required for text watermark",
                        new[] { nameof(WatermarkText) }));
                }

                if (string.IsNullOrWhiteSpace(FontName))
                {
                    results.Add(new ValidationResult(
                        "Font name is required for text watermark",
                        new[] { nameof(FontName) }));
                }

                if (FontSize <= 0 || FontSize > 144)
                {
                    results.Add(new ValidationResult(
                        "Font size must be between 1 and 144 points",
                        new[] { nameof(FontSize) }));
                }

                var validLayouts = new[] { WatermarkLayout.Diagonal, WatermarkLayout.Horizontal };
                if (!validLayouts.Contains(Layout))
                {
                    results.Add(new ValidationResult(
                        "Invalid layout",
                        new[] { nameof(Layout) }));
                }

                // Validate color format (hex color)
                if (!IsValidHexColor(Color as string))
                {
                    results.Add(new ValidationResult(
                        "Invalid color format. Use hex format like #FF0000",
                        new[] { nameof(Color) }));
                }
            }

            // Validate image watermark settings
            if (WatermarkType == WatermarkType.Image)
            {
                if (ScaleFactor < 0 || ScaleFactor > 65.5)
                {
                    results.Add(new ValidationResult(
                        "Scale factor must be between 0 (0%) and 65.5 (6550%)",
                        new[] { nameof(ScaleFactor) }));
                }
            }

            return results;
        }

        // Helper methods
        public bool IsTextWatermark()
        {
            return WatermarkType == WatermarkType.Text;
        }

        public bool IsImageWatermark()
        {
            return WatermarkType == WatermarkType.Image;
        }

        public double GetScaleFactorPercentage()
        {
            return ScaleFactor * 100;
        }

        public void SetScaleFactorFromPercentage(double percentage)
        {
            ScaleFactor = percentage / 100.0;
        }

        bool IsValidHexColor(string color)
        {
            if (string.IsNullOrWhiteSpace(color))
                return false;

            if (!color.StartsWith("#"))
                return false;

            if (color.Length != 7)
                return false;

            for (int i = 1; i < color.Length; i++)
            {
                char c = color[i];
                if (!(c >= '0' && c <= '9' || c >= 'A' && c <= 'F' || c >= 'a' && c <= 'f'))
                    return false;
            }

            return true;
        }
    }
}
