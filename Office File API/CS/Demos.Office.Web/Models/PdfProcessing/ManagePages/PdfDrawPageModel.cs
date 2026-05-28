using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.ManagePages {
    public class PdfDrawPageModel : PdfModelBase {
        public const string DefaultText = "Sample Text";
        //Translation - "This message was generated on {date}"
        public const string BidirectionalText = "تم إنشاء هذه الرسالة بتاريخ";
        [MaxLength(50, ErrorMessage = "Text must not exceed 50 symbols")]
        public string Text { get; set; } = DefaultText;
        public bool DrawBidirectional { get; set; }
        public bool DrawRectangle { get; set; } = true;
        public bool DrawEllipse { get; set; } = true;
        public bool DrawPolygon { get; set; } = true;
        public bool DrawLines { get; set; } = true;
        public bool DrawBeziers { get; set; } = true;
        public bool DrawPath { get; set; } = true;
        internal List<FormatInfo> OperationsTypes { get; } = new() {
            new() { Text = "Draw Text", Value = PdfDrawPageOperationType.Text },
            new() { Text = "Draw Image", Value = PdfDrawPageOperationType.Image },
            new() { Text = "Draw Shapes", Value = PdfDrawPageOperationType.Shapes }
        };
        public PdfDrawPageOperationType OperationType { get; set; } = PdfDrawPageOperationType.Text;
        public void ResetText() {
            Text = DrawBidirectional ? $"{DateTime.Today.ToShortDateString()} {BidirectionalText}" : DefaultText;
        }
    }

    public enum PdfDrawPageOperationType {
        Text, Image, Shapes
    }
}
