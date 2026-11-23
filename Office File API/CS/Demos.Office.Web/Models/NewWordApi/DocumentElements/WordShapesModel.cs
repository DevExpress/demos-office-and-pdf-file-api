using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.WordProcessing;
using DevExpress.Docs.Office;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentElements {
    public class WordShapesModel : WordFileModelBase {
        public override string OutputName { get; set; } = "Shape_Document";
        [Required(ErrorMessage = "Shape type is required")]
        public ShapeActionType ShapeType { get; set; } = ShapeActionType.Picture;

        [Range(0, 5, ErrorMessage = "Image scale must be between 0 and 5")]
        public float ImageScale { get; set; } = 1.0f;

        public GeometryPresetType ShapeGeometry { get; set; } = GeometryPresetType.Rectangle;
        public string ShapeText { get; set; } = "Sample Text";

        [Range(0, 6, ErrorMessage = "Outline width must be between 0 and 6 points")]
        public float OutlineWidth { get; set; } = 1.0f;

        public string OutlineColor { get; set; }
        internal object ShapeOutlineColor { get; set; } = "#9E67FF";

        public string FillColor { get; set; }
        internal object ShapeFillColor { get; set; } = "#FFFFFF";

        [Range(1, 5, ErrorMessage = "Number of shapes must be between 1 and 5")]
        public int NumberOfShapes { get; set; } = 3;

        internal bool IsFigureType => ShapeType is ShapeActionType.Figure or ShapeActionType.Group;
        internal bool IsGroupType => ShapeType == ShapeActionType.Group;
        internal bool IsTextboxType => ShapeType == ShapeActionType.Textbox;
        internal bool IsPictureType => ShapeType == ShapeActionType.Picture;

        internal readonly List<object> ShapeTypes = new() {
            new { Text = "Picture", Value = ShapeActionType.Picture },
            new { Text = "Textbox", Value = ShapeActionType.Textbox },
            new { Text = "Figure", Value = ShapeActionType.Figure },
            new { Text = "Shape Group", Value = ShapeActionType.Group }
        };
        internal readonly List<object> ShapeGeometries = new() {
            new { Text = "Rectangle", Value = GeometryPresetType.Rectangle },
            new { Text = "Circle", Value = GeometryPresetType.Oval },
            new { Text = "Star", Value = GeometryPresetType.Star4 },
            new { Text = "Heart", Value = GeometryPresetType.Heart },
            new { Text = "Arrow", Value = GeometryPresetType.RightArrow }
        };

        internal readonly List<FormatInfo> OutputFormats = new() {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };
    }
    public enum ShapeActionType {
        Picture,
        Textbox,
        Figure,
        Group
    }
}