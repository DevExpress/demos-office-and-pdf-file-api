using BlazorDemo.OfficeFileApi.Models.WordProcessing;

namespace BlazorDemo.OfficeFileApi.Models.AI.DocumentOperations {
    public class WordTranslationModel : WordFileModelBase {
        public override string OutputName { get; set; } = "ResultDocument";

        public string TargetLanguage { get; set; } = "de-DE";

        public readonly List<object> TargetLanguageList = new() {
            new { Text = "English", Value = "en-US" },
            new { Text = "French", Value = "fr-FR" },
            new { Text = "German", Value = "de-DE" },
            new { Text = "Spanish", Value = "es-ES" },
            new { Text = "Chinese", Value = "zh-CN" }
        };
    }
}
