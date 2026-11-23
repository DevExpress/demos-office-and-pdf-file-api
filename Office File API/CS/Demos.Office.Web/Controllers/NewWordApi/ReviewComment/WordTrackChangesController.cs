using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.ReviewComment;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs.Word;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.ReviewComment {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordTrackChangesController : WordControllerBase {
        [HttpGet("track-changes/template")]
        public IActionResult DownloadTemplate() {
            using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentTrackChangesTemplate, FileMode.Open, FileAccess.Read);
            var memoryStream = new MemoryStream();
            templateStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            return File(memoryStream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "TrackChanges_Template.docx");
        }

        [HttpPost("track-changes/process")]
        public async Task<IActionResult> ProcessTrackChangesTemplate(IFormFile file, [FromForm] WordTrackChangesModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Process track changes on the template
            var outputStream = ProcessTrackChangesDocument(inputStream, model.DisableTrackedChanges, model.AuthorFilter,
                model.RevisionAction == RevisionActionType.Accept, model.RevisionAction == RevisionActionType.Reject,
                NewWordFormatHelper.GetDocumentFormat(model.OutputFormat));
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ProcessTrackChangesDocument(Stream inputStream, bool disableTrackChanges,
                                           string authorFilter, bool acceptAll, bool rejectAll, DocumentFormat outputFormat) {
            // Load document
            using var document = new WordDocument(inputStream);

            if(acceptAll)
                AcceptRevisions(document, authorFilter);
            else if(rejectAll)
                RejectRevisions(document, authorFilter);

            document.TrackChanges = !disableTrackChanges;

            var outputStream = new MemoryStream();
            document.Save(outputStream, outputFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        void AcceptRevisions(WordDocument document, string authorFilter) {
            if(authorFilter == null)
                document.Revisions.AcceptAll();
            else
                document.Revisions.AcceptAll(revision => revision.Author.Contains(authorFilter, StringComparison.OrdinalIgnoreCase));
        }

        void RejectRevisions(WordDocument document, string authorFilter) {
            if(authorFilter == null)
                document.Revisions.RejectAll();
            else
                document.Revisions.RejectAll(revision => revision.Author.Contains(authorFilter, StringComparison.OrdinalIgnoreCase));
        }
        #endregion Code
    }
}