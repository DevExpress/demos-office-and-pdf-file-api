using BlazorDemo.OfficeFileApi.Models.WordProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.SearchExtract
{
    public class WordDocumentStatisticsModel : WordFileModelBase
    {
        public bool IncludeHeaderFooter { get; set; } = true;
        
        public bool IncludeTextBoxes { get; set; } = true;
        
        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf", ".txt" };
        public override string OutputFormat { get; set; } = "txt";
        public override string OutputName { get; set; } = "Document_Statistics";
    }
}