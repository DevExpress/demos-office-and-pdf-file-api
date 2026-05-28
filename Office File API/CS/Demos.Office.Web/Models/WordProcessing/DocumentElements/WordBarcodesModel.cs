using DevExpress.BarCodes;
using DevExpress.XtraRichEdit.API.Native;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements
{
    public class WordBarcodesModel : WordFileModelBase, IValidatableObject
    {

        public Symbology BarcodeType { get; set; } = Symbology.Code128;
        public string BarcodeValue { get; set; } = "123456789012";

        public override string OutputName { get; set; } = "Bracode_Document";

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            var results = new List<ValidationResult>();

            // Barcode validation
                if (string.IsNullOrWhiteSpace(BarcodeValue))
                {
                    results.Add(new ValidationResult("Barcode value is required", new[] { nameof(BarcodeValue) }));
                }
                else
                {
                    switch (BarcodeType)
                    {
                        case Symbology.Code128:
                            // Allow full ASCII 32-126. Limit length to a practical number (1-80)
                            if (BarcodeValue.Length > 80)
                                results.Add(new ValidationResult("Code128 value too long (max 80 chars)", new[] { nameof(BarcodeValue) }));
                            if (!BarcodeValue.All(c => c >= 32 && c <= 126))
                                results.Add(new ValidationResult("Code128 supports printable ASCII characters only", new[] { nameof(BarcodeValue) }));
                            break;
                        case Symbology.QRCode:
                            // Basic length guard
                            if (BarcodeValue.Length > 500)
                                results.Add(new ValidationResult("QR Code value too long (max 500 chars for this demo)", new[] { nameof(BarcodeValue) }));
                            // Alphanumeric mode allowed chars: 0-9 A-Z space $ % * + - . / :
                            // Force check against uppercase set. If user supplied lowercase, treat as invalid (caller can ToUpper() before submit if desired)
                            var qrPattern = "^[0-9A-Z $%*+\\-./:]+$";
                            if (!Regex.IsMatch(BarcodeValue, qrPattern))
                                results.Add(new ValidationResult("QR Code value must contain only digits 0-9, uppercase A-Z, space and $ % * + - . / : characters", new[] { nameof(BarcodeValue) }));
                            break;
                        case Symbology.DataBar:
                            // Numeric only, GS1 DataBar typically numeric (length limits simplified here)
                            if (!Regex.IsMatch(BarcodeValue, "^\\d{8,14}$"))
                                results.Add(new ValidationResult("Databar must be 8-14 digits", new[] { nameof(BarcodeValue) }));
                            break;
                        default:
                            results.Add(new ValidationResult("Unsupported barcode type", new[] { nameof(BarcodeType) }));
                            break;
                    }
                }

            return results;
        }

        internal readonly List<FormatInfo> OutputFormats = new()
        {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };

        internal readonly List<object> BarcodeTypes = new()
        {
            new { Text = "Code128", Value = Symbology.Code128 },
            new { Text = "QRCode", Value = Symbology.QRCode },
            new { Text = "Databar", Value = Symbology.DataBar }
        };
    }
}
