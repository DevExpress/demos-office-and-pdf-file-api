using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.Formatting;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs;
using DevExpress.Docs.Office;
using DevExpress.Docs.Word;
using System.Drawing;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.Formatting {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordTextFormattingController : WordControllerBase {
        [HttpPost("text-formatting/process")]
        public async Task<IActionResult> FormatWordDocuments(IFormFile file, [FromForm] WordTextFormattingModel model) {
            var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputStream = FormatDocument(inputStream, model);
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        Stream FormatDocument(Stream inputStream, WordTextFormattingModel model) {
            using var document = new WordDocument(inputStream);

            // Apply text formatting based on the model
            var textFragments = GetTextFragments(document, model.RangeType);
            if(model.IsApplyAction) {
                var foreColor = ColorTranslator.FromHtml(model.ForeColorHex ?? string.Empty);
                var backColor = ColorTranslator.FromHtml(model.BackColorHex ?? string.Empty);
                FormatText(textFragments, model.FontName, model.FontSize, foreColor, backColor,
                    model.Underline, model.IsBold, model.IsItalic);
            }
            else
                ResetCharacterFormatting(textFragments, model.CharacterResetValues ?? Array.Empty<TextFormattingAttribute>());

            var outputStream = new MemoryStream();
            if(model.OutputFormat == "pdf")
                document.ExportToPdf(outputStream);
            else
                document.Save(outputStream, NewWordFormatHelper.GetDocumentFormat(model.OutputFormat));
            outputStream.Position = 0;
            return outputStream;
        }

        #region Code
        List<TextFragment> GetTextFragments(WordDocument document, TextFormattingRangeType rangeType) {
            var result = new List<TextFragment>();
            if(rangeType == TextFormattingRangeType.FirstParagraph) {
                var firstParagraph = FindFirstParagraph(document.Body);
                if(firstParagraph != null)
                    CollectTextFragments(firstParagraph.InlineElements, result);
            }
            else if(rangeType == TextFormattingRangeType.FirstSection)
                CollectTextFragments(GetFirstSectionBlocks(document), result);
            else
                CollectTextFragments(document.Body, result);

            return result;
        }

        List<Block> GetFirstSectionBlocks(WordDocument document) {
            var result = new List<Block>();
            foreach(var block in document.Body) {
                result.Add(block);
                if(block is Paragraph paragraph && paragraph.Mark.Type != ParagraphMarkType.ParagraphMark)
                    break;
            }

            return result;
        }

        Paragraph FindFirstParagraph(IEnumerable<Block> blocks) {
            foreach(var block in blocks) {
                if(block is Paragraph paragraph)
                    return paragraph;
                if(block is Table table) {
                    foreach(var row in table.Rows) {
                        foreach(var cell in row.Cells) {
                            var cellParagraph = FindFirstParagraph(cell.Content);
                            if(cellParagraph != null)
                                return cellParagraph;
                        }
                    }
                }
            }

            return null;
        }

        void CollectTextFragments(IEnumerable<DocumentElement> elements, List<TextFragment> result) {
            foreach(var element in elements) {
                switch(element) {
                    case TextFragment textFragment:
                        result.Add(textFragment);
                        break;
                    case Paragraph paragraph:
                        CollectTextFragments(paragraph.InlineElements, result);
                        break;
                    case Hyperlink hyperlink:
                        CollectTextFragments(hyperlink.Content, result);
                        break;
                    case Field field:
                        CollectTextFragments(field.Result, result);
                        break;
                    case Table table: {
                        foreach(var row in table.Rows) {
                            foreach(var cell in row.Cells)
                                CollectTextFragments(cell.Content, result);
                        }

                        break;
                    }
                }
            }
        }

        void FormatText(IEnumerable<TextFragment> textFragments, string fontName, float fontSize,
                        Color foreColor, Color backColor, TextUnderlineType underline, bool bold, bool italic) {
            foreach(var textFragment in textFragments) {
                var properties = textFragment.CharacterProperties;
                properties.Font.Name = fontName;
                properties.FontSize = fontSize;
                properties.ForeColor = new OfficeColor(foreColor);
                properties.BackColor = new OfficeColor(backColor);
                properties.Underline = underline;
                properties.Bold = bold;
                properties.Italic = italic;
            }
        }

        void ResetCharacterFormatting(IEnumerable<TextFragment> textFragments, IEnumerable<TextFormattingAttribute> attributes) {
            foreach(var textFragment in textFragments) {
                var properties = textFragment.CharacterProperties;
                foreach(var attribute in attributes) {
                    switch(attribute) {
                        case TextFormattingAttribute.FontName:
                            properties.Font.Name = null;
                            break;
                        case TextFormattingAttribute.FontSize:
                            properties.FontSize = null;
                            break;
                        case TextFormattingAttribute.ForeColor:
                            properties.ForeColor = null;
                            break;
                        case TextFormattingAttribute.BackColor:
                            properties.BackColor = null;
                            break;
                        case TextFormattingAttribute.Bold:
                            properties.Bold = null;
                            break;
                        case TextFormattingAttribute.Italic:
                            properties.Italic = null;
                            break;
                        case TextFormattingAttribute.Underline:
                            properties.Underline = null;
                            break;
                    }
                }
            }
        }
        #endregion Code
    }
}