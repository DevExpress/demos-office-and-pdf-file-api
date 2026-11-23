using BlazorDemo.OfficeFileApi.Models.WordProcessing;
using DevExpress.Docs.Word;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentElements {
    public class WordListsModel : WordFileModelBase, IValidatableObject {
        [Required(ErrorMessage = "List type is required")]
        public NumberingType ListType { get; set; } = NumberingType.Bullet;

        [Range(1, 20, ErrorMessage = "Number of list items must be between 1 and 20")]
        public int NumberOfItems { get; set; } = 5;

        [Required(ErrorMessage = "Bullet style is required")]
        public BulletedStyleType BulletStyle { get; set; } = BulletedStyleType.Disc;

        public string CustomBulletText { get; set; } = "•";

        [Required(ErrorMessage = "Numbering format is required")]
        public NumberingFormat NumberingFormat { get; set; } = NumberingFormat.Decimal;

        [Range(1, 1000, ErrorMessage = "Start number must be between 1 and 1000")]
        public int StartNumber { get; set; } = 1;

        [Range(1, 3, ErrorMessage = "Maximum levels must be between 1 and 3")]
        public int MaxLevels { get; set; } = 3;

        [Required(ErrorMessage = "Multilevel style is required")]
        public MultilevelStyleType MultilevelStyle { get; set; } = MultilevelStyleType.Outline;

        [Range(0, 5, ErrorMessage = "Indent size must be between 0 and 5 inches")]
        public float IndentSize { get; set; } = 0.5f;

        [Range(0, 2, ErrorMessage = "Hanging indent must be between 0 and 2 inches")]
        public float HangingIndent { get; set; } = 0.25f;

        [Range(0, 72, ErrorMessage = "Space before must be between 0 and 72 points")]
        public float SpaceBefore { get; set; } = 0f;

        [Range(0, 72, ErrorMessage = "Space after must be between 0 and 72 points")]
        public float SpaceAfter { get; set; } = 0f;

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "List_Result";

        internal readonly List<object> ListTypes = new() {
            new { Text = "Bullet List", Value = NumberingType.Bullet },
            new { Text = "Numbered List", Value = NumberingType.Simple },
            new { Text = "Multilevel List", Value = NumberingType.MultiLevel }
        };

        internal readonly List<object> BulletStyles = new() {
            new { Text = "Disc (●)", Value = BulletedStyleType.Disc },
            new { Text = "Circle (○)", Value = BulletedStyleType.Circle },
            new { Text = "Square (■)", Value = BulletedStyleType.Square },
            new { Text = "Arrow (→)", Value = BulletedStyleType.Arrow },
            new { Text = "Check (✓)", Value = BulletedStyleType.Check },
            new { Text = "Custom", Value = BulletedStyleType.Custom }
        };

        internal readonly List<object> NumberingFormats = new() {
            new { Text = "1, 2, 3 (Arabic)", Value = NumberingFormat.Decimal },
            new { Text = "I, II, III (Roman Upper)", Value = NumberingFormat.UpperRoman },
            new { Text = "i, ii, iii (Roman Lower)", Value = NumberingFormat.LowerRoman },
            new { Text = "A, B, C (Letter Upper)", Value = NumberingFormat.UpperLetter },
            new { Text = "a, b, c (Letter Lower)", Value = NumberingFormat.LowerLetter }
        };

        internal readonly List<object> MultilevelStyles = new() {
            new { Text = "Outline (1., 1.1, 1.1.1)", Value = MultilevelStyleType.Outline },
            new { Text = "Legal (1., a., i.)", Value = MultilevelStyleType.Legal },
            new { Text = "Mixed Bullet (●, ○, ■)", Value = MultilevelStyleType.Bullet }
        };

        internal readonly List<FormatInfo> OutputFormats = new() {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };
        public string GetBulletCharacter(BulletedStyleType bulletStyle, string customText) {
            return bulletStyle switch {
                BulletedStyleType.Disc => "●",
                BulletedStyleType.Circle => "○",
                BulletedStyleType.Square => "■",
                BulletedStyleType.Arrow => "→",
                BulletedStyleType.Check => "✓",
                BulletedStyleType.Custom => customText ?? "•",
                _ => "●"
            };
        }
        public List<string> GenerateMultilevelDisplayFormats(MultilevelStyleType style) {
            return style switch {
                MultilevelStyleType.Legal => new List<string> {
                    "%1.",
                    "%2.",
                    "%3."
                },
                MultilevelStyleType.Bullet => new List<string> {
                    "●",
                    "○",
                    "■"
                },
                _ => new List<string> {
                    "%1.",
                    "%1.%2",
                    "%1.%2.%3"
                },
            };
        }
        public List<NumberingFormat> GenerateMultilevelNumberingFormats(MultilevelStyleType style) {
            return style switch {
                MultilevelStyleType.Legal => new List<NumberingFormat> {
                    NumberingFormat.Decimal,
                    NumberingFormat.LowerLetter,
                    NumberingFormat.LowerRoman
                },
                MultilevelStyleType.Bullet => new List<NumberingFormat> {
                    NumberingFormat.Bullet,
                    NumberingFormat.Bullet,
                    NumberingFormat.Bullet
                },
                _ => new List<NumberingFormat> {
                    NumberingFormat.Decimal,
                    NumberingFormat.Decimal,
                    NumberingFormat.Decimal
                },
            };
        }
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            var results = new List<ValidationResult>();

            if (NumberOfItems <= 0 || NumberOfItems > 50) {
                results.Add(new ValidationResult(
                    "Number of list items must be between 1 and 50",
                    new[] { nameof(NumberOfItems) }));
            }

            if (ListType == NumberingType.Bullet) {
                if (BulletStyle == BulletedStyleType.Custom && string.IsNullOrWhiteSpace(CustomBulletText)) {
                    results.Add(new ValidationResult(
                        "Custom bullet text is required when using custom bullet style",
                        new[] { nameof(CustomBulletText) }));
                }
            }

            if (ListType == NumberingType.Simple) {
                var validNumberingFormats = new[] { NumberingFormat.Decimal, NumberingFormat.UpperRoman, NumberingFormat.LowerRoman, NumberingFormat.UpperLetter, NumberingFormat.LowerLetter };
                if (!validNumberingFormats.Contains(NumberingFormat)) {
                    results.Add(new ValidationResult(
                        "Invalid numbering format",
                        new[] { nameof(NumberingFormat) }));
                }

                if (StartNumber <= 0 || StartNumber > 1000) {
                    results.Add(new ValidationResult(
                        "Start number must be between 1 and 1000",
                        new[] { nameof(StartNumber) }));
                }
            }

            if (ListType == NumberingType.MultiLevel) {
                if (MaxLevels <= 0 || MaxLevels > 9) {
                    results.Add(new ValidationResult(
                        "Maximum levels must be between 1 and 9",
                        new[] { nameof(MaxLevels) }));
                }
            }

            if (IndentSize < 0 || IndentSize > 5) {
                results.Add(new ValidationResult(
                    "Indent size must be between 0 and 5 inches",
                    new[] { nameof(IndentSize) }));
            }

            if (HangingIndent < 0 || HangingIndent > 2) {
                results.Add(new ValidationResult(
                    "Hanging indent must be between 0 and 2 inches",
                    new[] { nameof(HangingIndent) }));
            }

            return results;
        }

        public bool IsBulletList() {
            return ListType == NumberingType.Bullet;
        }

        public bool IsNumberingList() {
            return ListType == NumberingType.Simple;
        }

        public bool IsMultilevelList() {
            return ListType == NumberingType.MultiLevel;
        }

        public bool IsCustomBullet() {
            return ListType == NumberingType.Bullet && BulletStyle == BulletedStyleType.Custom;
        }
    }
    public enum NumberingType {
        Bullet,
        Simple,
        MultiLevel
    }
    public enum MultilevelStyleType {
        Outline,
        Legal,
        Bullet
    }
    public enum BulletedStyleType {
        Disc,
        Circle,
        Square,
        Arrow,
        Check,
        Custom
    }
}