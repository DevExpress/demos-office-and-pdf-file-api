using System.ComponentModel.DataAnnotations;
using DevExpress.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.Formatting {
    public class ExcelCellFormattingModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Cell_Formatting";

        public bool CreateCustomStyle { get; set; } = false;
        public string FontName { get; set; } = "Segoe UI";
        public float FontSize { get; set; } = 16f;
        public SpreadsheetFontStyle FontStyle { get; set; } = SpreadsheetFontStyle.Regular;
        public string ForeColorHex { get; set; }
        internal object ForeColor { get; set; } = "#000000";
        public string BackColorHex { get; set; }
        internal object BackColor { get; set; } = "#66CDFF";
        public SpreadsheetHorizontalAlignment HorizontalAlignment { get; set; } = SpreadsheetHorizontalAlignment.General;
        public SpreadsheetVerticalAlignment VerticalAlignment { get; set; } = SpreadsheetVerticalAlignment.Center;     

        internal readonly List<object> FontNames = new()
        {
            new { Text = "Calibri", Value = "Calibri" },
            new { Text = "Segoe UI", Value = "Segoe UI" },
            new { Text = "Arial", Value = "Arial" },
            new { Text = "Times New Roman", Value = "Times New Roman" },
            new { Text = "Verdana", Value = "Verdana" },
            new { Text = "Tahoma", Value = "Tahoma" },
            new { Text = "Georgia", Value = "Georgia" },
            new { Text = "Impact", Value = "Impact" }
        };
        internal readonly List<object> FontStyles = new()
        {
            new { Text = "Regular", Value = SpreadsheetFontStyle.Regular },
            new { Text = "Italic", Value = SpreadsheetFontStyle.Italic },
            new { Text = "Bold", Value = SpreadsheetFontStyle.Bold },
            new { Text = "Bold Italic", Value = SpreadsheetFontStyle.BoldItalic }
        };
        internal readonly List<object> HorizontalAlignments = new(){
            new { Text = "Center", Value = SpreadsheetHorizontalAlignment.Center },
            new { Text = "CenterContinuous", Value = SpreadsheetHorizontalAlignment.CenterContinuous },
            new { Text = "Distributed", Value = SpreadsheetHorizontalAlignment.Distributed },
            new { Text = "Fill", Value = SpreadsheetHorizontalAlignment.Fill },
            new { Text = "General", Value = SpreadsheetHorizontalAlignment.General },
            new { Text = "Justify", Value = SpreadsheetHorizontalAlignment.Justify },
            new { Text = "Left", Value = SpreadsheetHorizontalAlignment.Left },
            new { Text = "Right", Value = SpreadsheetHorizontalAlignment.Right },
        };
        internal readonly List<object> VerticalAlignments = new(){
            new { Text = "Bottom", Value = SpreadsheetVerticalAlignment.Bottom },
            new { Text = "Center", Value = SpreadsheetVerticalAlignment.Center },
            new { Text = "Distributed", Value = SpreadsheetVerticalAlignment.Distributed },
            new { Text = "Justify", Value = SpreadsheetVerticalAlignment.Justify },
            new { Text = "Top", Value = SpreadsheetVerticalAlignment.Top },
        };
    }
}
