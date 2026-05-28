using BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.DocumentElements {
    [ApiController]
    [Route("api/wordprocessing/lists")]
    public class WordListsController : WordControllerBase {
        [HttpPost("generate")]
        public async Task<IActionResult> GenerateListDocument([FromForm] WordListsModel model) {
            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate document with list
            Stream outputStream = new MemoryStream();
            switch (model.ListType) {
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
        Stream CreateBulletList(int itemCount, string bullet, float indent, float hangingIndent, float spacingBeforePoints, float spacingAfterPoints, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            var document = wordProcessor.Document;

            // Set document unit to inches for precise sizing
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            for (int i = 1; i <= itemCount; i++) {
                var itemText = $"Item {i}\n";
                document.AppendText(itemText);
            }
            document.Delete(document.Paragraphs.Last().Range);

            // Create a new list pattern object
            AbstractNumberingList list = document.AbstractNumberingLists.Add();

            //Specify the list's type
            list.NumberingType = NumberingType.Bullet;
            ListLevel level = list.Levels[0];
            ApplyListLevelFormatting(level, NumberingFormat.Bullet, bullet, 1, indent, hangingIndent, spacingBeforePoints, spacingAfterPoints);

            //Create a new list based on the specific pattern
            NumberingList bulletedList = document.NumberingLists.Add(0);

            // Add paragraphs to the list
            ParagraphCollection paragraphs = document.Paragraphs;
            paragraphs.AddParagraphsToList(document.Range, bulletedList, 0);

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        Stream CreateNumberingList(int itemCount, NumberingFormat numberingFormat, int startNumber, float indent,
            float hangingIndent, float spacingBeforePoints, float spacingAfterPoints, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            var document = wordProcessor.Document;

            // Set document unit to inches for precise sizing
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            for (int i = 1; i <= itemCount; i++) {
                var itemText = $"Item {i}\n";
                document.AppendText(itemText);
            }
            document.Delete(document.Paragraphs.Last().Range);

            // Create a new list pattern object
            AbstractNumberingList list = document.AbstractNumberingLists.Add();

            //Specify the list's type
            list.NumberingType = NumberingType.Simple;
            ListLevel level = list.Levels[0];
            ApplyListLevelFormatting(level, numberingFormat, "{0}", startNumber, indent, hangingIndent, spacingBeforePoints, spacingAfterPoints);

            //Create a new list based on the specific pattern
            NumberingList numberingList = document.NumberingLists.Add(0);

            // Add paragraphs to the list
            ParagraphCollection paragraphs = document.Paragraphs;
            paragraphs.AddParagraphsToList(document.Range, numberingList, 0);

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        Stream CreateMultilevelList(int itemCount, int levels, List<NumberingFormat> numberingFormats, List<string> displayFormats, float indent,
            float hangingIndent, float spacingBeforePoints, float spacingAfterPoints, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            var document = wordProcessor.Document;

            // Set document unit to inches for precise sizing
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            for (int i = 1; i <= itemCount; i++) {
                for (int j = 1; j <= levels; j++) {
                    var itemText = $"Item {i} - Level {j}\n";
                    document.AppendText(itemText);
                }
            }
            document.Delete(document.Paragraphs.Last().Range);

            // Create a new list pattern object
            AbstractNumberingList list = document.AbstractNumberingLists.Add();

            //Specify the list's type
            list.NumberingType = NumberingType.MultiLevel;

            for (int i = 0; i < levels; i++) {
                var level = list.Levels[i];
                ApplyListLevelFormatting(level, numberingFormats[i], displayFormats[i], 1, indent + 0.1f * i, hangingIndent + 0.1f * i, spacingBeforePoints, spacingAfterPoints);
            }

            //Create a new list based on the specific pattern
            NumberingList multilevelList = document.NumberingLists.Add(0);

            // Add paragraphs to the list
            ParagraphCollection paragraphs = document.Paragraphs;
            paragraphs.AddParagraphsToList(document.Range, multilevelList, 0);
            foreach (Paragraph paragraph in paragraphs)
                paragraph.ListLevel = paragraph.Index % levels;

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        void ApplyListLevelFormatting(ListLevel level, NumberingFormat numberingFormat, string displayFormat, int startNumber, float indent,
            float hangingIndent, float spacingBeforePoints, float spacingAfterPoints) {
            level.NumberingFormat = numberingFormat;
            level.DisplayFormatString = displayFormat;
            level.Start = startNumber;

            level.ParagraphProperties.LeftIndent = indent + hangingIndent;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = hangingIndent;

            //Convert points to inches for spacing
            level.ParagraphProperties.SpacingBefore = spacingBeforePoints / 72;
            level.ParagraphProperties.SpacingAfter = spacingAfterPoints / 72;
        }
        #endregion Code
    }
}
