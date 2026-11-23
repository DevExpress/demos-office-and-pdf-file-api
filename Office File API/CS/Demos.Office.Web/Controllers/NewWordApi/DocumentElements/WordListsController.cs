using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentElements;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs.Word;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.DocumentElements {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordListsController : WordControllerBase {
        [HttpPost("lists/generate")]
        public IActionResult GenerateListDocument([FromForm] WordListsModel model) {
            // Determine target format
            var targetFormat = NewWordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate document with list
            Stream outputStream = new MemoryStream();
            switch(model.ListType) {
                case NumberingType.Bullet:
                    outputStream = CreateBulletList(model.NumberOfItems, model.GetBulletCharacter(model.BulletStyle, model.CustomBulletText ?? string.Empty), model.IndentSize,
                        model.HangingIndent, model.SpaceBefore, model.SpaceAfter, targetFormat);
                    break;
                case NumberingType.Simple:
                    outputStream = CreateNumberingList(model.NumberOfItems, model.NumberingFormat, model.StartNumber, model.IndentSize,
                        model.HangingIndent, model.SpaceBefore, model.SpaceAfter, targetFormat);
                    break;
                case NumberingType.MultiLevel:
                    outputStream = CreateMultilevelList(model.NumberOfItems, model.MaxLevels, model.GenerateMultilevelNumberingFormats(model.MultilevelStyle),
                        model.GenerateMultilevelDisplayFormats(model.MultilevelStyle), model.IndentSize, model.HangingIndent, model.SpaceBefore, model.SpaceAfter, targetFormat);
                    break;
            }

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        const float DocumentUnitsPerInch = 300f;
        const float DocumentUnitsPerPoint = DocumentUnitsPerInch / 72f;

        Stream CreateBulletList(int itemCount, string bullet, float indent, float hangingIndent, float spacingBeforePoints, float spacingAfterPoints, DocumentFormat targetFormat) {
            using var document = new WordDocument();

            // Create a new list pattern object
            var template = new ListTemplate();
            ApplyListLevelFormatting(template.Levels[0], NumberingFormat.Bullet, bullet, 1, indent, hangingIndent, spacingBeforePoints, spacingAfterPoints);
            // Create a new list based on the specific pattern
            var bulletedList = document.NumberingLists.Add(template);

            // Add paragraphs to the list
            for(var i = 1; i <= itemCount; i++) {
                var paragraph = document.Body.AddParagraph($"Item {i}");
                paragraph.NumberingList = bulletedList;
                paragraph.ListLevelIndex = 0;
            }

            return CreateOutputStream(document, targetFormat);
        }

        Stream CreateNumberingList(int itemCount, NumberingFormat numberingFormat, int startNumber, float indent,
                                   float hangingIndent, float spacingBeforePoints, float spacingAfterPoints, DocumentFormat targetFormat) {
            using var document = new WordDocument();

            // Create a new list pattern object
            var template = new ListTemplate();
            ApplyListLevelFormatting(template.Levels[0], numberingFormat, "%1", startNumber, indent, hangingIndent, spacingBeforePoints, spacingAfterPoints);
            // Create a new list based on the specific pattern
            var numberingList = document.NumberingLists.Add(template);

            // Add paragraphs to the list
            for(var i = 1; i <= itemCount; i++) {
                var paragraph = document.Body.AddParagraph($"Item {i}");
                paragraph.NumberingList = numberingList;
                paragraph.ListLevelIndex = 0;
            }

            return CreateOutputStream(document, targetFormat);
        }

        Stream CreateMultilevelList(int itemCount, int levels, List<NumberingFormat> numberingFormats, List<string> displayFormats, float indent,
                                    float hangingIndent, float spacingBeforePoints, float spacingAfterPoints, DocumentFormat targetFormat) {
            using var document = new WordDocument();

            // Create a new list pattern object
            var template = new ListTemplate();
            for(var i = 0; i < levels; i++)
                ApplyListLevelFormatting(template.Levels[i], numberingFormats[i], displayFormats[i], 1, indent + 0.1f * i, hangingIndent + 0.1f * i, spacingBeforePoints, spacingAfterPoints);
            // Create a new list based on the specific pattern
            var multilevelList = document.NumberingLists.Add(template);

            // Add paragraphs to the list
            for(var i = 1; i <= itemCount; i++) {
                for(var j = 1; j <= levels; j++) {
                    var paragraph = document.Body.AddParagraph($"Item {i} - Level {j}");
                    paragraph.NumberingList = multilevelList;
                    paragraph.ListLevelIndex = j - 1;
                }
            }

            return CreateOutputStream(document, targetFormat);
        }

        void ApplyListLevelFormatting(ListLevel level, NumberingFormat numberingFormat, string displayFormat, int startNumber, float indent,
                                      float hangingIndent, float spacingBeforePoints, float spacingAfterPoints) {
            level.Format = numberingFormat;
            level.DisplayFormat = displayFormat;
            level.Start = startNumber;

            level.ParagraphProperties.LeftIndent = (indent + hangingIndent) * DocumentUnitsPerInch;
            level.ParagraphProperties.FirstLineIndent = -hangingIndent * DocumentUnitsPerInch;

            // Convert points to document units for spacing
            level.ParagraphProperties.SpacingBefore = spacingBeforePoints * DocumentUnitsPerPoint;
            level.ParagraphProperties.SpacingAfter = spacingAfterPoints * DocumentUnitsPerPoint;
        }

        Stream CreateOutputStream(WordDocument document, DocumentFormat targetFormat) {
            var outputStream = new MemoryStream();
            document.Save(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}