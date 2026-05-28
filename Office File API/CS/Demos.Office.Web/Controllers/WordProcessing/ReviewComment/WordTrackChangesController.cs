using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.ReviewComment;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.ReviewComment {
    [ApiController]
    [Route("api/wordprocessing/track-changes")]
    public class WordTrackChangesController : WordControllerBase {
        [HttpGet("template")]
        public IActionResult DownloadTemplate() {
            using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentTrackChangesTemplate, FileMode.Open, FileAccess.Read);
            var memoryStream = new MemoryStream();
            templateStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            return File(memoryStream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "TrackChanges_Template.docx");
        }

        [HttpPost("process")]
        public async Task<IActionResult> ProcessTrackChangesTemplate(IFormFile file, [FromForm] WordTrackChangesModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Process track changes on the template
            var outputStream = ProcessTrackChangesDocument(inputStream, model.DisableTrackedChanges, model.AuthorFilter, model.RevisionAction == RevisionActionType.Accept, model.RevisionAction == RevisionActionType.Reject, targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ProcessTrackChangesDocument(Stream inputStream, bool disableTrackChanges,
            string authorFilter, bool acceptAll, bool rejectAll, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();

            // Load document
            wordProcessor.LoadDocument(inputStream);
            var document = wordProcessor.Document;

            if (acceptAll)
                AcceptRevisions(document, authorFilter);
            else if (rejectAll)
                RejectRevisions(document, authorFilter);

            document.TrackChanges.Enabled = !disableTrackChanges;

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, outputFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        void AcceptRevisions(Document document, string authorFilter) {
            if (authorFilter == null)
                document.Revisions.AcceptAll();
            else
                document.Revisions.AcceptAll(revision => revision.Author.Contains(authorFilter, StringComparison.OrdinalIgnoreCase));
        }

        void RejectRevisions(Document document, string authorFilter) {
            if (authorFilter == null)
                document.Revisions.RejectAll();
            else
                document.Revisions.RejectAll(revision => revision.Author.Contains(authorFilter, StringComparison.OrdinalIgnoreCase));
        }
        #endregion Code
    }
}
