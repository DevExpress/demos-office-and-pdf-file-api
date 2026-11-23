using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.SearchExtract
{
    public class WordMetadataModel : WordFileModelBase
    {
        public bool ReadDocumentProperties { get; set; } = true;
        
        public bool ReadCustomProperties { get; set; } = true;
        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "Document_Metadata";
        public override string OutputFormat { get; set; }= "txt"; 
    }
}