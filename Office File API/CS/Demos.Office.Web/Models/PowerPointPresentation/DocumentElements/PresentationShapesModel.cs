using System.ComponentModel.DataAnnotations;
using DevExpress.Docs.Office;

namespace BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentElements {
    public class PresentationShapesModel : PresentationModelBase {
        public override string OutputName { get; set; } = "Presentation_Shapes";

        public ShapeActionType ShapeActionType { get; set; } = ShapeActionType.Picture;
        public GeometryPresetType ShapeGeometryPresetType { get; set; } = GeometryPresetType.Rectangle;

        public string ShapeText { get; set; } = "Sample Text";

        [Range(0, 6, ErrorMessage = "Outline width must be between 0 and 6 points")]
        public float OutlineWidth { get; set; } = 1.0f; // points

        public string OutlineColor { get; set; }
        internal object ShapeOutlineColor { get; set; } = "#9E67FF";

        public string FillColor { get; set; }
        internal object ShapeFillColor { get; set; } = "#FFFFFF";

        [Range(1, 5, ErrorMessage = "Number of shapes must be between 1 and 5")]
        public int NumberOfShapes { get; set; } = 3;

        internal bool IsFigureType => ShapeActionType is ShapeActionType.Figure or ShapeActionType.Group;
        internal bool IsGroupType => ShapeActionType == ShapeActionType.Group;
        internal bool IsTextboxType => ShapeActionType == ShapeActionType.Textbox;
        internal bool IsPictureType => ShapeActionType == ShapeActionType.Picture;

        internal readonly List<object> ShapeActionTypes = new()
        {
            new { Text = "Picture", Value = ShapeActionType.Picture },
            new { Text = "Textbox", Value = ShapeActionType.Textbox },
            new { Text = "Figure", Value = ShapeActionType.Figure },
            new { Text = "Shape Group", Value = ShapeActionType.Group }
        };
        internal readonly List<object> ShapeGeometryPresetTypes = new()
        {
            new { Text = "Rectangle", Value = GeometryPresetType.Rectangle },
            new { Text = "Circle", Value = GeometryPresetType.Oval },
            new { Text = "Star", Value = GeometryPresetType.Star4 },
            new { Text = "Heart", Value = GeometryPresetType.Heart },
            new { Text = "Arrow", Value = GeometryPresetType.RightArrow }
        };
        public List<FormatInfo> OutputFormats { get; set; } = new() {
            new FormatInfo { Text = "PowerPoint Presentation (.pptx)", Value = "pptx" },
            new FormatInfo { Text = "PowerPoint Macro-Enabled Presentation (.pptm)", Value = "pptm" },
            new FormatInfo { Text = "PowerPoint Template (.potx)", Value = "potx" },
            new FormatInfo { Text = "PowerPoint Macro-Enabled Template (.potm)", Value = "potm" }
        };
    }
    public enum ShapeActionType {
        Picture,
        Textbox,
        Figure,
        Group
    }
}
