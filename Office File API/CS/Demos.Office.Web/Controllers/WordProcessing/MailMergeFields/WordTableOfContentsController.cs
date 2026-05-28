using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.MailMergeFields;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.MailMergeFields {
    [ApiController]
    [Route("api/wordprocessing/toc")]
    public class WordTableOfContentsController : WordControllerBase {
        [HttpGet("template")]
        public IActionResult DownloadTemplate() {
            using var templateStream = new FileStream(DemoDocumentPaths.WordTableOfContentsTemplateFile, FileMode.Open, FileAccess.Read);
            var memoryStream = new MemoryStream();
            templateStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            return File(memoryStream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "TableOfContents_Template.docx");
        }

        [HttpPost("process")]
        public async Task<IActionResult> CreateTableOfContentsDocument([FromForm] WordTableOfContentsModel model) {
            using var templateStream = new FileStream(DemoDocumentPaths.WordTableOfContentsTemplateFile, FileMode.Open, FileAccess.Read);
            Stream stream = new MemoryStream();
            templateStream.CopyTo(stream);
            stream.Position = 0;

            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);
            var outputStream = CreateTableOfContents(stream, model.TabLeaderType, model.IncludePageNumbers, model.InsertMode, model.StartLevel, model.EndLevel, model.UseHyperlinks, targetFormat);
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        MemoryStream CreateTableOfContents(Stream stream, TabLeaderType tabLeaderType, bool includePageNumbers,
            TOCInsertMode insertMode, int startLevel, int endLevel, bool useHyperlinks, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(stream);

            List<string> switches = new();

            // Hyperlinks switch
            if (useHyperlinks)
                switches.Add("\\h");

            // Levels (only if not TC fields selection and valid range)
            if (insertMode != TOCInsertMode.TCFields)
                    switches.Add($"\\o \"{startLevel}-{endLevel}\"");

            // Show / hide page numbers (if IncludePageNumbers false add \n to hide)
            if (!includePageNumbers)
                switches.Add($"\\n \"{startLevel}-{endLevel}\"");
            // OutlineLevels vs ParagraphStyles vs TC fields
            switch (insertMode) {
                case TOCInsertMode.OutlineLevels:
                    switches.Add("\\u"); // use outline levels
                    break;
                case TOCInsertMode.TCFields:
                    // Use identifier for TC fields (choose static id) \f id
                    switches.Add($"\\f chapters");
                    switches.Add($"\\l \"{startLevel}-{endLevel}\"");
                    break;
            }
            string fieldCode = "TOC" + (switches.Count > 0 ? (" " + string.Join(" ", switches)) : string.Empty) + " ";

            var document = wordProcessor.Document;
            var tocParagraphRange = document.Bookmarks["toc_bookmark"].Range;
            Field tocField = document.Fields.Create(tocParagraphRange.Start, fieldCode);
            tocField.Update();
            document.UpdateAllFields();

            var output = new MemoryStream();
            wordProcessor.SaveDocument(output, targetFormat);
            output.Position = 0;
            return output;
        }
        #endregion Code
    }
}
