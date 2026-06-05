using BlazorDemo.OfficeFileApi.Models.WordProcessing;

namespace BlazorDemo.OfficeFileApi.Models.AI.DocumentOperations {
    public class WordProofreadingModel : WordFileModelBase {
        public override string OutputName { get; set; } = "ResultDocument";

        public string SourceLanguage { get; set; } = "en-US";

        public readonly List<object> SourceLanguageList = new() {
            new { Text = "English", Value = "en-US" },
            new { Text = "French", Value = "fr-FR" },
            new { Text = "German", Value = "de-DE" },
            new { Text = "Spanish", Value = "es-ES" },
            new { Text = "Chinese", Value = "zh-CN" }
        };
    }
}
