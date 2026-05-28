using System.ComponentModel.DataAnnotations;
using DevExpress.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    public class ExcelShapesModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Shape_Document";
        [Required(ErrorMessage = "Shape type is required")]
        public ShapeActionType ShapeType { get; set; } = ShapeActionType.Picture;

        public bool Embedded { get; set; } = true;

        public ShapeGeometryPreset ShapeGeometry { get; set; } = ShapeGeometryPreset.Rectangle;
        public string ShapeText { get; set; } = "Sample Text";

        [Range(0, 6, ErrorMessage = "Outline width must be between 0 and 6 points")]
        public float OutlineWidth { get; set; } = 1.0f; // points

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

        internal readonly List<object> ShapeTypes = new()
        {
            new { Text = "Picture", Value = ShapeActionType.Picture },
            new { Text = "Textbox", Value = ShapeActionType.Textbox },
            new { Text = "Figure", Value = ShapeActionType.Figure },
            new { Text = "Shape Group", Value = ShapeActionType.Group }
        };
        internal readonly List<object> ShapeGeometries = new()
        {
            new { Text = "Rectangle", Value = ShapeGeometryPreset.Rectangle },
            new { Text = "Circle", Value = ShapeGeometryPreset.Oval },
            new { Text = "Star", Value = ShapeGeometryPreset.Star4 },
            new { Text = "Heart", Value = ShapeGeometryPreset.Heart },
            new { Text = "Arrow", Value = ShapeGeometryPreset.RightArrow }
        };

        public override List<FormatInfo> OutputFormats { get; set; } = new() {
            new FormatInfo { Text = "Excel Workbook (*.xlsx)", Value = "xlsx" },
            new FormatInfo { Text = "Excel Macro-Enabled Workbook (*.xlsm)", Value = "xlsm" },
            new FormatInfo { Text = "Excel Template (*.xltx)", Value = "xltx" },
            new FormatInfo { Text = "Excel 97-2003 Template (*.xlt)", Value = "xlt" },
            new FormatInfo { Text = "Excel Macro-Enabled Template (*.xltm)", Value = "xltm" },
            new FormatInfo { Text = "Excel Binary Workbook (*.xlsb)", Value = "xlsb" }
        };
    }
    public enum ShapeActionType {
        Picture,
        Textbox,
        Figure,
        Group
    }
}
