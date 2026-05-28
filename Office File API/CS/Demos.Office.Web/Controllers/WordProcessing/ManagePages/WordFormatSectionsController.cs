using BlazorDemo.OfficeFileApi.Models.WordProcessing.ManagePages;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.Drawing.Printing;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.ManagePages {
    [ApiController]
    [Route("api/wordprocessing/format-sections")]
    public class WordFormatSectionsController : WordControllerBase {
        [HttpPost("process")]
        public async Task<IActionResult> FormatSections(IFormFile file, [FromForm] WordFormatSectionsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Format document sections
            var formattedStream = FormatDocumentSections(inputStream, model, targetFormat);

            return CreateOutputFile(formattedStream, model.OutputFormat, model.OutputName);
        }
        Stream FormatDocumentSections(Stream inputStream, WordFormatSectionsModel model, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();

            // Load document
            wordProcessor.LoadDocument(inputStream);
            var document = wordProcessor.Document;

            // Apply formatting to all sections
            foreach (Section section in document.Sections) {
                ApplySectionFormatting(document, section, model);
            }

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, outputFormat);
            outputStream.Position = 0;
            return outputStream;
        }
        void ApplySectionFormatting(Document document, Section section, WordFormatSectionsModel model) {
            // Apply paper size and orientation
            SetPaperSizeAndOrientation(document, section, model.PaperSize,
                model.CustomPaperWidth, model.CustomPaperHeight, model.IsLandscape);

            // Apply margins
            SetMargins(document, section, model.LeftMargin, model.RightMargin, model.TopMargin, model.BottomMargin, model.MirroredMargins);

            // Apply header and footer offsets
            SetHeaderFooterOffsets(document, section, model.HeaderOffset, model.FooterOffset);

            // Apply gutter settings
            SetGutterSettings(document, section, GutterHelper.IsGutterAtTop(model.GutterPosition), GutterHelper.GetGutterPosition(model.GutterPosition), model.GutterSize);
        }
        #region Code
        void SetPaperSizeAndOrientation(Document document, Section section, DXPaperKind paperKind, float customPageWidth, float customPageHeight, bool isLandscape) {
            document.Unit = DevExpress.Office.DocumentUnit.Inch;
            section.Page.PaperKind = paperKind;
            if (paperKind == DXPaperKind.Custom) {
                section.Page.Width = customPageWidth;
                section.Page.Height = customPageHeight;
            }
            section.Page.Landscape = isLandscape;
        }

        void SetMargins(Document document, Section section, float leftMargin, float rightMargin, float topMargin, float bottomMargin, bool mirroredMargins) {
            document.Unit = DevExpress.Office.DocumentUnit.Inch;
            section.Margins.Top = topMargin;
            section.Margins.Bottom = bottomMargin;
            section.Margins.Left = leftMargin;
            section.Margins.Right = rightMargin;
            document.MarginType = mirroredMargins ? MarginType.Mirrored : MarginType.Normal;
        }

        void SetHeaderFooterOffsets(Document document, Section section, float headerOffset, float footerOffset) {
            document.Unit = DevExpress.Office.DocumentUnit.Inch;
            section.Margins.HeaderOffset = headerOffset;
            section.Margins.FooterOffset = footerOffset;
        }

        void SetGutterSettings(Document document, Section section, bool isGutterAtTop, GutterPosition gutterPosition, float gutter) {
            document.Unit = DevExpress.Office.DocumentUnit.Inch;
            section.Margins.GutterPosition = gutterPosition;
            section.Margins.Gutter = gutter;
            document.GutterAtTop = isGutterAtTop;
        }
        #endregion Code
    }
}
