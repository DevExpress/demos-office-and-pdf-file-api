using BlazorDemo.OfficeFileApi.Models.WordProcessing.Formatting;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.Formatting {
    [ApiController]
    [Route("api/wordprocessing")]
    public class WordParagraphFormattingController : WordControllerBase {
        [HttpPost("paragraph-formatting/process")]
        public async Task<IActionResult> FormatWordDocuments(IFormFile file, [FromForm] WordParagraphFormattingModel model) {
            var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            Stream outputStream = FormatDocument(inputStream, model, targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        Stream FormatDocument(Stream inputStream, WordParagraphFormattingModel model, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.Unit = DevExpress.Office.DocumentUnit.Inch;
            wordProcessor.LoadDocument(inputStream);
            Document document = wordProcessor.Document;
            DocumentRange range = null;
            if (model.RangeType == FormattedRangeType.FirstParagraph)
                range = document.Paragraphs[0].Range;
            else if (model.RangeType == FormattedRangeType.FirstPage) {
                FixedRange fixedRange = wordProcessor.DocumentLayout.GetPage(0).MainContentRange;
                range = document.CreateRange(fixedRange.Start, fixedRange.Length);
            }
            else
                range = document.Range;
            if (model.IsApplyAction) {
                FormatParagraph(document, range, ColorTranslator.FromHtml(model.ShadingHex ?? string.Empty), model.ParagraphAlignment, model.LineSpacingMultiplier, model.LeftIndent, model.WidowOrphanControl, model.KeepWithNext, model.KeepLinesTogether);
                if (model.ApplyParagraphBorders)
                    FormatParagraphBorders(document, range, model.BorderLineWidth, model.BorderLineStyle, ColorTranslator.FromHtml(model.BorderColorHex ?? string.Empty));
            }
            else
                ResetParagraphFormatting(document, range, GetParagraphPropertiesMask(model.ParagraphResetValues));
            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;

        }

        ParagraphPropertiesMask GetParagraphPropertiesMask(IEnumerable<ParagraphPropertiesMask> paragraphResetValues) {
            ParagraphPropertiesMask mask = ParagraphPropertiesMask.None;
            if (paragraphResetValues != null)
                foreach (var value in paragraphResetValues)
                            mask |= value;
            return mask;
        }
        #region Code
        void FormatParagraph(Document document, DocumentRange range, Color shading, ParagraphAlignment alignment,
            float lineSpacingMultiplier, float leftIndent, bool widowOrphanControl, bool keepWithNext, bool KeepLinesTogether) {
            ParagraphProperties pp = document.BeginUpdateParagraphs(range);
            pp.BackColor = shading;
            pp.Alignment = alignment;
            pp.LineSpacingType = ParagraphLineSpacing.Multiple;
            pp.LineSpacingMultiplier = lineSpacingMultiplier;
            pp.LeftIndent = leftIndent;
            pp.WidowOrphanControl = widowOrphanControl;
            pp.KeepWithNext = keepWithNext;
            pp.KeepLinesTogether = KeepLinesTogether;

            document.EndUpdateParagraphs(pp);
        }

        void ResetParagraphFormatting(Document document, DocumentRange range, ParagraphPropertiesMask resetParagraphProperties) {
            ParagraphProperties cp = document.BeginUpdateParagraphs(range);
            cp.Reset(resetParagraphProperties);
            document.EndUpdateParagraphs(cp);
        }

        void FormatParagraphBorders(Document document, DocumentRange range, float lineWidth, BorderLineStyle lineStyle, Color borderColor) {
            ParagraphProperties pp = document.BeginUpdateParagraphs(range);
            SetParagraphBorder(pp.Borders.HorizontalBorder, lineWidth, lineStyle, borderColor);
            SetParagraphBorder(pp.Borders.BottomBorder, lineWidth, lineStyle, borderColor);
            SetParagraphBorder(pp.Borders.TopBorder, lineWidth, lineStyle, borderColor);
            SetParagraphBorder(pp.Borders.LeftBorder, lineWidth, lineStyle, borderColor);
            SetParagraphBorder(pp.Borders.RightBorder, lineWidth, lineStyle, borderColor);
            document.EndUpdateParagraphs(pp);
        }

        void SetParagraphBorder(ParagraphBorder border, float lineWidth, BorderLineStyle lineStyle, Color borderColor) {
            border.LineWidth = lineWidth;
            border.LineStyle = lineStyle;
            border.LineColor = borderColor;
        }
        #endregion Code
    }
}
