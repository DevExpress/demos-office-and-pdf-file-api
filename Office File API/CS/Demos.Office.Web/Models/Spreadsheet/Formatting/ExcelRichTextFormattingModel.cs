using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.Formatting {
    public class ExcelRichTextFormattingModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Rich_Text_Formatting";

        [Required(ErrorMessage = "First text part is required.")]
        [MinLength(3)]
        [MaxLength(10)]
        public string FirstTextPart { get; set; } = "Rich";
        [Required(ErrorMessage = "Second text part is required.")]
        [MinLength(3)]
        [MaxLength(10)]
        public string SecondTextPart { get; set; } = "Text";
        [Required(ErrorMessage = "Third text part is required.")]
        [MinLength(3)]
        [MaxLength(10)]
        public string ThirdTextPart { get; set; } = "Formatting";

        public string FirstPartColorHex { get; set; }
        internal object FirstPartColor { get; set; } = "#C59FC9";

        public string SecondPartColorHex { get; set; }
        internal object SecondPartColor { get; set; } = "#2C608E";

        public string ThirdPartColorHex { get; set; }
        internal object ThirdPartColor { get; set; } = "#2F244F";
    }
}
