using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.ManagePages {
    public class PdfClearPagesModel : PdfPageRangeModelBase {
        public bool ClearAnnotations { get; set; } = true;
        public bool ClearGraphics    { get; set; } = true;
        public bool ClearImages      { get; set; } = true;
        public bool ClearText        { get; set; } = true;
    }
}
