using DevExpress.AIIntegration;

namespace BlazorDemo.OfficeFileApi.Models.AI.DocumentOperations {
    public class PresentationSummarizationModel : PowerPointPresentation.PresentationModelBase {
        public override string OutputName { get; set; } = "ResultPresentation";

        public SummarizationMode SummarizationMode { get; set; } = SummarizationMode.Extractive;
        public OutputMode OutputMode { get; set; } = OutputMode.InsertAboveContent;

        public readonly List<object> SummarizationModeList = new() {
            new { Text = "Extractive", Value = SummarizationMode.Extractive },
            new { Text = "Abstractive", Value = SummarizationMode.Abstractive }
        };
        public readonly List<object> OutputModeList = new() {
            new { Text = "Beginning of Presentation", Value = OutputMode.InsertAboveContent },
            new { Text = "End of Presentation", Value = OutputMode.InsertBelowContent },
            new { Text = "Standalone File (TXT)", Value = OutputMode.StandaloneFile }
        };
    }
}
