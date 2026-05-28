using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.MergeSplit {
    public class PdfSplitDocumentsModel : PdfPageRangeModelBase {
        public override string OutputFormat { get => "zip"; }
        public override string OutputName   { get => "Documents"; }
    }
}
