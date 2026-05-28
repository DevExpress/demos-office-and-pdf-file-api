using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.DocumentProtection {
    public class PdfApplyWatermarkModel : PdfModelBase {
        [MaxLength(15, ErrorMessage = "Text must not exceed 15 symbols")]
        public string WatermarkText { get; set; } = "WATERMARK";
        public bool Rotate { get; set; } = false;
    }
}
