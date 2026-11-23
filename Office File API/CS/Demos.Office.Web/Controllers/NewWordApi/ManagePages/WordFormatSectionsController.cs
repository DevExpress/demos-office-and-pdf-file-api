using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.ManagePages;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs.Word;
using DevExpress.Drawing.Printing;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.ManagePages {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordFormatSectionsController : WordControllerBase {
        const float DocumentUnitsPerInch = 300f;

        [HttpPost("format-sections/process")]
        public async Task<IActionResult> FormatSections(IFormFile file, [FromForm] WordFormatSectionsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Format document sections
            var formattedStream = FormatDocumentSections(inputStream, model);
            return CreateOutputFile(formattedStream, model.OutputFormat, model.OutputName);
        }
        Stream FormatDocumentSections(Stream inputStream, WordFormatSectionsModel model) {
            // Load document
            using var document = new WordDocument(inputStream);

            // Apply formatting to all sections
            foreach(var section in document.Sections) {
                ApplySectionFormatting(document, section, model);
            }

            var outputStream = new MemoryStream();
            document.Save(outputStream, NewWordFormatHelper.GetDocumentFormat(model.OutputFormat));
            outputStream.Position = 0;
            return outputStream;
        }
        void ApplySectionFormatting(WordDocument document, Section section, WordFormatSectionsModel model) {
            // Apply paper size and orientation
            var orientation = model.IsLandscape ? PageOrientation.Landscape : PageOrientation.Portrait;
            SetPaperSizeAndOrientation(section, model.PaperSize,
                ToDocumentUnits(model.CustomPaperWidth), ToDocumentUnits(model.CustomPaperHeight), orientation);

            // Apply margins
            SetMargins(document, section, ToDocumentUnits(model.LeftMargin), ToDocumentUnits(model.RightMargin),
                ToDocumentUnits(model.TopMargin), ToDocumentUnits(model.BottomMargin), model.MirroredMargins);

            // Apply header and footer offsets
            SetHeaderFooterOffsets(section, ToDocumentUnits(model.HeaderOffset), ToDocumentUnits(model.FooterOffset));

            // Apply gutter settings
            SetGutterSettings(document, section, IsGutterAtTop(model.GutterPosition), GetGutterPosition(model.GutterPosition), ToDocumentUnits(model.GutterSize));
        }
        float ToDocumentUnits(float inches) {
            return inches * DocumentUnitsPerInch;
        }
        bool IsGutterAtTop(string gutterPosition) {
            return gutterPosition.ToLowerInvariant() == "top";
        }
        GutterPosition GetGutterPosition(string gutterPosition) {
            return gutterPosition.ToLowerInvariant() == "right" ? GutterPosition.Right : GutterPosition.Left;
        }

        #region Code
        void SetPaperSizeAndOrientation(Section section, DXPaperKind paperKind, float customPageWidth, float customPageHeight,
                                        PageOrientation orientation) {
            if(paperKind == DXPaperKind.Custom) {
                section.Page.Orientation = PageOrientation.Portrait;
                section.Page.Width = customPageWidth;
                section.Page.Height = customPageHeight;
            }
            else {
                section.Page.PaperKind = paperKind;
            }

            section.Page.Orientation = orientation;
        }

        void SetMargins(WordDocument document, Section section, float leftMargin, float rightMargin, float topMargin, float bottomMargin, bool mirroredMargins) {
            section.Margins.Top = topMargin;
            section.Margins.Bottom = bottomMargin;
            section.Margins.Left = leftMargin;
            section.Margins.Right = rightMargin;
            document.MirrorMargins = mirroredMargins;
        }

        void SetHeaderFooterOffsets(Section section, float headerOffset, float footerOffset) {
            section.Margins.HeaderOffset = headerOffset;
            section.Margins.FooterOffset = footerOffset;
        }

        void SetGutterSettings(WordDocument document, Section section, bool isGutterAtTop, GutterPosition gutterPosition, float gutter) {
            section.Margins.GutterPosition = gutterPosition;
            section.Margins.Gutter = gutter;
            document.GutterAtTop = isGutterAtTop;
        }
        #endregion Code
    }
}