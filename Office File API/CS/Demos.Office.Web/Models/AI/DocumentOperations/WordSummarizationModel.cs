using BlazorDemo.OfficeFileApi.Models.WordProcessing;
using DevExpress.AIIntegration;

namespace BlazorDemo.OfficeFileApi.Models.AI.DocumentOperations {
    public class WordSummarizationModel : WordFileModelBase {
        public override string OutputName { get; set; } = "ResultDocument";

        public SummarizationMode SummarizationMode { get; set; } = SummarizationMode.Extractive;
        public OutputMode OutputMode { get; set; } = OutputMode.InsertAboveContent;

        public readonly List<object> SummarizationModeList = new() {
            new { Text = "Extractive", Value = SummarizationMode.Extractive },
            new { Text = "Abstractive", Value = SummarizationMode.Abstractive }
        };
        public readonly List<object> OutputModeList = new() {
            new { Text = "Insert Above Content", Value = OutputMode.InsertAboveContent },
            new { Text = "Insert Below Content", Value = OutputMode.InsertBelowContent },
            new { Text = "Standalone File (TXT)", Value = OutputMode.StandaloneFile }
        };
    }

    public enum OutputMode { InsertAboveContent, InsertBelowContent, StandaloneFile }
}
