using BlazorDemo.OfficeFileApi.Models.PdfProcessing;
using DevExpress.AIIntegration;

namespace BlazorDemo.OfficeFileApi.Models.AI.DocumentOperations {
    public class PdfSummarizationModel : PdfModelBase {
        public override string OutputFormat { get => "txt"; }
        public override string OutputName { get; set; } = "ResultDocument";

        public SummarizationMode SummarizationMode { get; set; } = SummarizationMode.Extractive;

        public readonly List<object> SummarizationModeList = new() {
            new { Text = "Extractive", Value = SummarizationMode.Extractive },
            new { Text = "Abstractive", Value = SummarizationMode.Abstractive }
        };
    }
}
