using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.SearchExtract;

#region Code

using DevExpress.Docs.Word;
using System.Text;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.SearchExtract {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordDocumentStatisticsController : WordControllerBase {
        [HttpPost("document-statistics/analyze")]
        public async Task<IActionResult> AnalyzeDocumentStatistics(IFormFile file, [FromForm] WordDocumentStatisticsModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Generate plain text statistics output
            var output = new StringBuilder();
            var statisticsText = GenerateDocumentStatisticsText(inputStream, model.IncludeHeaderFooter, model.IncludeTextBoxes);
            output.AppendLine(statisticsText);

            using var outputStream = new MemoryStream();
            using(var writer = new StreamWriter(outputStream, Encoding.UTF8)) {
                writer.Write(output.ToString());
                writer.Flush();
                outputStream.Position = 0;
                return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
            }
        }

        #region Code
        string GenerateDocumentStatisticsText(Stream inputStream, bool includeHeaderFooter, bool includeTextBoxes) {
            var output = new StringBuilder();
            using var document = new WordDocument(inputStream);

            // Basic document statistics
            output.AppendLine("BASIC DOCUMENT STATISTICS:");
            output.AppendLine(new string('-', 35));

            // Count paragraphs, words, characters, tables, images, hyperlinks, bookmarks
            var contentParts = new List<IEnumerable<Block>> { document.Body };
            if(includeHeaderFooter) {
                foreach(var section in document.Sections) {
                    contentParts.AddRange(section.Headers);
                    contentParts.AddRange(section.Footers);
                }
            }

            var paragraphs = new List<Paragraph>();
            var tables = new List<Table>();
            foreach(var contentPart in contentParts.Distinct())
                CollectBlocks(contentPart, paragraphs, tables);
            var inlineElements = new List<Inline>();
            CollectInlineElements(paragraphs, inlineElements);
            if(includeTextBoxes) {
                var textBoxShapes = new List<Shape>();
                CollectTextBoxShapes(inlineElements.OfType<ShapeBase>(), textBoxShapes);
                var textBoxParagraphs = new List<Paragraph>();
                foreach(var shape in textBoxShapes) {
                    if(shape.Content != null)
                        CollectBlocks(shape.Content, textBoxParagraphs, tables);
                }

                paragraphs.AddRange(textBoxParagraphs);
                CollectInlineElements(textBoxParagraphs, inlineElements);
            }

            var paragraphTexts = paragraphs.Select(paragraph => paragraph.GetText()).ToList();
            var totalWords = paragraphTexts.Sum(CountWords);
            var totalChars = paragraphTexts.Sum(text => text.Length);
            var totalCharsNoSpaces = paragraphTexts.Sum(text => text.Count(character => !char.IsWhiteSpace(character)));

            output.AppendLine($"Sections: {document.Sections.Count}");
            output.AppendLine($"Paragraphs: {paragraphs.Count}");
            output.AppendLine();

            // Text statistics
            output.AppendLine("TEXT CONTENT ANALYSIS:");
            output.AppendLine(new string('-', 35));

            output.AppendLine($"Words: {totalWords:N0}");
            output.AppendLine($"Characters (with spaces): {totalChars:N0}");
            output.AppendLine($"Characters (no spaces): {totalCharsNoSpaces:N0}");
            output.AppendLine();

            // Document elements count
            output.AppendLine("DOCUMENT ELEMENTS:");
            output.AppendLine(new string('-', 35));
            output.AppendLine($"Tables: {tables.Count:N0}");
            output.AppendLine($"Images/Shapes: {inlineElements.OfType<ShapeBase>().Count():N0}");
            output.AppendLine($"Hyperlinks: {inlineElements.OfType<Hyperlink>().Count():N0}");
            output.AppendLine($"Bookmarks: {CountBookmarks(document, includeTextBoxes):N0}");
            output.AppendLine($"Comments: {document.Comments.Count:N0}");
            output.AppendLine();

            return output.ToString();
        }

        void CollectBlocks(IEnumerable<Block> blocks, List<Paragraph> paragraphs, List<Table> tables) {
            foreach(var block in blocks) {
                if(block is Paragraph paragraph)
                    paragraphs.Add(paragraph);
                else if(block is Table table) {
                    tables.Add(table);
                    foreach(var row in table.Rows) {
                        foreach(var cell in row.Cells)
                            CollectBlocks(cell.Content, paragraphs, tables);
                    }
                }
            }
        }

        void CollectInlineElements(IEnumerable<DocumentElement> elements, List<Inline> result) {
            foreach(var element in elements) {
                if(element is Paragraph paragraph)
                    CollectInlineElements(paragraph.InlineElements, result);
                else if(element is Inline inline) {
                    result.Add(inline);
                    if(inline is Hyperlink hyperlink)
                        CollectInlineElements(hyperlink.Content, result);
                    else if(inline is Field field)
                        CollectInlineElements(field.Result, result);
                }
            }
        }

        void CollectTextBoxShapes(IEnumerable<ShapeBase> shapes, List<Shape> result) {
            foreach(var shape in shapes) {
                if(shape is Shape textBoxShape)
                    result.Add(textBoxShape);
                else if(shape is ContainerShapeBase containerShape)
                    CollectTextBoxShapes(containerShape.Shapes, result);
            }
        }

        int CountBookmarks(WordDocument document, bool includeTextBoxes) {
            var result = document.Bookmarks.Count(bookmark =>
                bookmark.DocumentFragment == document.Body
                || (includeTextBoxes && bookmark.DocumentFragment is ShapeTextBoxContent));
            return result;
        }

        int CountWords(string text) {
            if(string.IsNullOrWhiteSpace(text))
                return 0;

            return text.Split(new char[] { ' ', '\r', '\n', '\t', '\v' }, StringSplitOptions.RemoveEmptyEntries).Length;
        }
        #endregion Code
    }
}