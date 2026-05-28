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
    public class WordTextFormattingController : WordControllerBase {
        [HttpPost("text-formatting/process")]
        public async Task<IActionResult> FormatWordDocuments(IFormFile file, [FromForm] WordTextFormattingModel model) {
            var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            Stream outputStream = FormatDocument(inputStream, model, targetFormat);
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        Stream FormatDocument(Stream inputStream, WordTextFormattingModel model, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();

            wordProcessor.LoadDocument(inputStream);
            Document document = wordProcessor.Document;
            // Apply text formatting based on the model
            DocumentRange range = null;
            if (model.RangeType == FormattedRangeType.FirstParagraph)
                range = document.Paragraphs[0].Range;
            else if (model.RangeType == FormattedRangeType.FirstPage) {
                FixedRange fixedRange = wordProcessor.DocumentLayout.GetPage(0).MainContentRange;
                range = document.CreateRange(fixedRange.Start, fixedRange.Length);
            }
            else
                range = document.Range;
            if (model.IsApplyAction)
                FormatText(document, range, model.FontName, model.FontSize, ColorTranslator.FromHtml(model.ForeColorHex ?? string.Empty), ColorTranslator.FromHtml(model.BackColorHex ?? string.Empty),
                   model.Underline, model.IsBold, model.IsItalic);
            else
                ResetCharacterFormatting(document, range, GetCharacterPropertiesMask(model.CharacterResetValues));
            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;

        }
        CharacterPropertiesMask GetCharacterPropertiesMask(IEnumerable<CharacterPropertiesMask> resetValues) {
            CharacterPropertiesMask mask = CharacterPropertiesMask.None;
            if (resetValues != null) {
                foreach (var value in resetValues) {
                    mask |= value;
                }
            }
            return mask;
        }
        #region Code
        void FormatText(Document document, DocumentRange range, string fontName, float fontSize,
            Color foreColor, Color backgroundColor, UnderlineType underline, bool bold, bool italic) {
            CharacterProperties cp = document.BeginUpdateCharacters(range);
            cp.FontName = fontName;
            cp.FontSize = fontSize;
            cp.ForeColor = foreColor;
            cp.BackColor = backgroundColor;
            cp.Underline = underline;
            cp.Bold = bold;
            cp.Italic = italic;

            document.EndUpdateCharacters(cp);
        }

        void ResetCharacterFormatting(Document document, DocumentRange range, CharacterPropertiesMask resetCharacterProperties) {
            CharacterProperties cp = document.BeginUpdateCharacters(range);
            cp.Reset(resetCharacterProperties);
            document.EndUpdateCharacters(cp);
        }
        #endregion Code
    }
}
