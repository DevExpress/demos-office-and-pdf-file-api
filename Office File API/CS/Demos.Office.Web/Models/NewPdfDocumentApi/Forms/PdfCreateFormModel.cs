using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.Forms {
    public class PdfCreateFormModel : PdfModelBase {
        public bool FillFromFile { get; set; } = true;
    }
}
