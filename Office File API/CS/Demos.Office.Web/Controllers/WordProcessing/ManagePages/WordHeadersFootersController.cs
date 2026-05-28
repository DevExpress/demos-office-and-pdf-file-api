using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.ManagePages;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.ManagePages {
    [ApiController]
    [Route("api/wordprocessing/headers-footers")]
    public class WordHeadersFootersController : WordControllerBase {
        [HttpGet("template")]
        public IActionResult DownloadTemplate() {
            try
            {
                using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentHeaderFooterTemplate, FileMode.Open, FileAccess.Read);
                var memoryStream = new MemoryStream();
                templateStream.CopyTo(memoryStream);
                memoryStream.Position = 0;

                return File(memoryStream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "HeaderFooter_Template.docx");
            }
            catch (Exception ex) {
                return StatusCode(500, new { Error = "Error downloading template", Details = ex.Message });
            }
        }

        [HttpPost("process")]
        public async Task<IActionResult> ProcessHeadersFootersTemplate([FromForm] WordHeadersFootersModel model) {
            using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentHeaderFooterTemplate, FileMode.Open, FileAccess.Read);
            Stream inputStream = new MemoryStream();
            templateStream.CopyTo(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Process headers and footers on the template
            var outputStream = ProcessHeadersFootersDocument(templateStream, model, targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        Stream ProcessHeadersFootersDocument(FileStream templateStream, WordHeadersFootersModel model, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(templateStream);
            Document document = wordProcessor.Document;

            document.Sections[0].DifferentFirstPage = model.DifferentFirstPage;
            if (model.DifferentFirstPage) {
                AddHeader(wordProcessor, HeaderFooterType.First, model.FirstPageHeaderText, model.FirstPageHeaderAlignment);
                AddFooter(wordProcessor, HeaderFooterType.First, model.FirstPageFooterText, model.FirstPageFooterAlignment, model.AddPageNumbering);
            }
            if (model.AddHeader)
                AddHeader(wordProcessor, HeaderFooterType.Primary, model.HeaderText, model.HeaderAlignment);
            if (model.AddFooter)
                AddFooter(wordProcessor, HeaderFooterType.Primary, model.FooterText, model.FooterAlignment, model.AddPageNumbering);

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        #region Code
        void AddHeader(RichEditDocumentServer wordProcessor, HeaderFooterType headerFooterType, string text, ParagraphAlignment alignment) {
            Document document = wordProcessor.Document;
            foreach (Section section in document.Sections) {
                SubDocument header = section.BeginUpdateHeader(headerFooterType);
                header.InsertText(header.Paragraphs[0].Range.Start, text ?? "");
                ParagraphProperties props = header.BeginUpdateParagraphs(header.Range);
                props.Alignment = alignment;
                header.EndUpdateParagraphs(props);
                section.EndUpdateHeader(header);
            }
        }

        void AddFooter(RichEditDocumentServer wordProcessor, HeaderFooterType headerFooterType, string text, ParagraphAlignment alignment, bool addPageNumbering) {
            Document document = wordProcessor.Document;
            foreach (Section section in document.Sections) {
                SubDocument footer = section.BeginUpdateFooter(headerFooterType);
                footer.InsertText(footer.Paragraphs[0].Range.Start, text ?? "");
                if (addPageNumbering) {
                    footer.Paragraphs.Append();
                    footer.AppendText("Page ");
                    footer.Fields.Create(footer.Range.End, "PAGE");
                    footer.AppendText(" of ");
                    footer.Fields.Create(footer.Range.End, "NUMPAGES");
                }
                ParagraphProperties props = footer.BeginUpdateParagraphs(footer.Range);
                props.Alignment = alignment;
                footer.EndUpdateParagraphs(props);
                section.EndUpdateFooter(footer);
            }
        }
        #endregion Code
    }
}
