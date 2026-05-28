using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements
{
    public class WordOleObjectsModel : WordFileModelBase
    {
        public override string OutputName { get; set; } = "OLE_Document";
        internal readonly List<FormatInfo> OutputFormats = new()
        {
            new FormatInfo{ Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" }
        };
        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".docx", ".docm", ".dotx", ".dotm" };
    }
}
