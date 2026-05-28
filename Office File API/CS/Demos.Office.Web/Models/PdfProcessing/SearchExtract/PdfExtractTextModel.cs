namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.SearchExtract {
    public class PdfExtractTextModel : PdfPageRangeModelBase {
        public override string OutputFormat { get => "txt"; }
        public override string OutputName { get => "Document Text"; }
    }
}
