using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.SearchExtract;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Text;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.SearchExtract {
    [ApiController]
    [Route("api/wordprocessing")]
    public class WordDocumentStatisticsController : WordControllerBase {
        [HttpPost("document-statistics/analyze")]
        public async Task<IActionResult> AnalyzeDocumentStatistics(IFormFile file, [FromForm] WordDocumentStatisticsModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Generate plain text statistics output
            var output = new StringBuilder();
            var statisticsText = GenerateDocumentStatisticsText(inputStream, model.IncludeHeaderFooter, model.IncludeTextBoxes, model.IncludeFootnotesEndnotes);
            output.AppendLine(statisticsText);

            using var outputStream = new MemoryStream();
            using(StreamWriter writer = new StreamWriter(outputStream, Encoding.UTF8)) {
                writer.Write(output.ToString());
                writer.Flush();
                outputStream.Position = 0;
                return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
            }
        }

        #region Code
        string GenerateDocumentStatisticsText(Stream inputStream, bool includeHeaderFooter, bool includeTextBoxes, bool includeFootnotesEndnotes) {
            var output = new StringBuilder();
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(inputStream);

            var document = wordProcessor.Document;

            // Basic document statistics
            output.AppendLine("BASIC DOCUMENT STATISTICS:");
            output.AppendLine(new string('-', 35));

            var pageCount = wordProcessor.DocumentLayout.GetPageCount();
            var sectionCount = document.Sections.Count;
            var paragraphCount = 0;
            var totalWords = 0;
            var totalChars = 0;
            var totalCharsNoSpaces = 0;
            var totalTables = 0;
            var totalImages = 0;
            var totalHyperlinks = 0;
            var totalBookmarks = 0;

            // Count paragraphs, words, characters, tables, images, hyperlinks, bookmarks
            wordProcessor.Document.ForEachSubDocument(subDoc => {
                var docType = subDoc.GetSubDocumentType();
                if((docType == SubDocumentType.Main) ||
                (includeHeaderFooter && (docType == SubDocumentType.Header || docType == SubDocumentType.Footer)) ||
                (includeFootnotesEndnotes && (docType == SubDocumentType.FootNote || docType == SubDocumentType.EndNote)) ||
                (includeTextBoxes && (docType == SubDocumentType.TextBox))) {
                    paragraphCount += subDoc.Paragraphs.Count;
                    string docText = subDoc.GetText(subDoc.Range);
                    totalWords += CountWords(docText);
                    totalChars += docText.Length;
                    totalCharsNoSpaces += docText.Replace(" ", "").Replace("\r", "").Replace("\n", "").Replace("\t", "").Length;
                    totalTables += subDoc.Tables.Count;
                    totalImages += subDoc.Shapes.Count;
                    totalHyperlinks += subDoc.Hyperlinks.Count;
                    totalBookmarks += subDoc.Bookmarks.Count;
                }
            });

            output.AppendLine($"Pages: {pageCount}");
            output.AppendLine($"Sections: {sectionCount}");
            output.AppendLine($"Paragraphs: {paragraphCount}");
            output.AppendLine();

            // Text statistics
            output.AppendLine("TEXT CONTENT ANALYSIS:");
            output.AppendLine(new string('-', 35));

            output.AppendLine($"Words: {totalWords:N0}");
            output.AppendLine($"Characters (with spaces): {totalChars:N0}");
            output.AppendLine($"Characters (no spaces): {totalCharsNoSpaces:N0}");
            output.AppendLine();

            // Document elements count
            output.AppendLine("DOCUMENT ELEMENTS:");
            output.AppendLine(new string('-', 35));
            output.AppendLine($"Tables: {totalTables:N0}");
            output.AppendLine($"Images/Shapes: {totalImages:N0}");
            output.AppendLine($"Hyperlinks: {totalHyperlinks:N0}");
            output.AppendLine($"Bookmarks: {totalBookmarks:N0}");
            output.AppendLine($"Comments: {document.Comments.Count:N0}");
            output.AppendLine($"Footnotes: {document.Footnotes.Count:N0}");
            output.AppendLine($"Endnotes: {document.Endnotes.Count:N0}");
            output.AppendLine();

            return output.ToString();
        }
        int CountWords(string text) {
            if(string.IsNullOrWhiteSpace(text))
                return 0;

            return text.Split(new char[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
        }
        #endregion Code
    }
}
