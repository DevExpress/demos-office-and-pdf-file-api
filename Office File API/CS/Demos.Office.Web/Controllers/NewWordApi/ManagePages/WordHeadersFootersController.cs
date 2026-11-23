using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.ManagePages;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs.Word;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.ManagePages {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordHeadersFootersController : WordControllerBase {
        [HttpGet("headers-footers/template")]
        public IActionResult DownloadTemplate() {
            try {
                using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentHeaderFooterTemplate, FileMode.Open, FileAccess.Read);
                var memoryStream = new MemoryStream();
                templateStream.CopyTo(memoryStream);
                memoryStream.Position = 0;

                return File(memoryStream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "HeaderFooter_Template.docx");
            }
            catch(Exception ex) {
                return StatusCode(500, new { Error = "Error downloading template", Details = ex.Message });
            }
        }

        [HttpPost("headers-footers/process")]
        public IActionResult ProcessHeadersFootersTemplate([FromForm] WordHeadersFootersModel model) {
            using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentHeaderFooterTemplate, FileMode.Open, FileAccess.Read);
            // Process headers and footers on the template
            var outputStream = ProcessHeadersFootersDocument(templateStream, model);
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        Stream ProcessHeadersFootersDocument(Stream templateStream, WordHeadersFootersModel model) {
            using var document = new WordDocument(templateStream);

            document.Sections[0].DifferentFirstPage = model.DifferentFirstPage;
            if(model.DifferentFirstPage) {
                AddHeader(document, HeaderFooterType.First, model.FirstPageHeaderText, model.FirstPageHeaderAlignment);
                AddFooter(document, HeaderFooterType.First, model.FirstPageFooterText, model.FirstPageFooterAlignment, false);
            }

            if(model.AddHeader)
                AddHeader(document, HeaderFooterType.Odd, model.HeaderText, model.HeaderAlignment);
            if(model.AddFooter)
                AddFooter(document, HeaderFooterType.Odd, model.FooterText, model.FooterAlignment, model.AddPageNumbering);

            var outputStream = new MemoryStream();
            document.Save(outputStream, NewWordFormatHelper.GetDocumentFormat(model.OutputFormat));
            outputStream.Position = 0;
            return outputStream;
        }

        #region Code
        void AddHeader(WordDocument document, HeaderFooterType headerFooterType, string text, ParagraphAlignment alignment) {
            foreach(var section in document.Sections) {
                var header = section.Headers[headerFooterType];
                if(header == null) {
                    header = new Header();
                    section.Headers[headerFooterType] = header;
                }

                if(!string.IsNullOrEmpty(text))
                    header.InsertText(0, text);
                SetAlignment(header, alignment);
            }
        }

        void AddFooter(WordDocument document, HeaderFooterType headerFooterType, string text, ParagraphAlignment alignment, bool addPageNumbering) {
            foreach(var section in document.Sections) {
                var footer = section.Footers[headerFooterType];
                if(footer == null) {
                    footer = new Footer();
                    section.Footers[headerFooterType] = footer;
                }

                if(!string.IsNullOrEmpty(text))
                    footer.InsertText(0, text);
                if(addPageNumbering) {
                    var paragraph = footer.AddParagraph("Page ");
                    paragraph.InlineElements.Add(new Field("PAGE"));
                    paragraph.AddText(" of ");
                    paragraph.InlineElements.Add(new Field("NUMPAGES"));
                }

                SetAlignment(footer, alignment);
            }
        }

        void SetAlignment(SectionHeaderFooterBase headerFooter, ParagraphAlignment alignment) {
            foreach(var paragraph in headerFooter.OfType<Paragraph>())
                paragraph.Properties.Alignment = alignment;
        }
        #endregion Code
    }
}