using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.GenerateDocument {
    public class PdfGenerateInvoiceModel : PdfModelBase {
        public bool IncludeZugferd { get; set; } = true;
    }
}
