using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.Formatting;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs;
using DevExpress.Docs.Word;
using System.Drawing;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.Formatting {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordParagraphFormattingController : WordControllerBase {
        const float DocumentUnitsPerInch = 300f;
        const float DocumentUnitsPerPoint = DocumentUnitsPerInch / 72f;

        [HttpPost("paragraph-formatting/process")]
        public async Task<IActionResult> FormatWordDocuments(IFormFile file, [FromForm] WordParagraphFormattingModel model) {
            var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputStream = FormatDocument(inputStream, model);
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        Stream FormatDocument(Stream inputStream, WordParagraphFormattingModel model) {
            using var document = new WordDocument(inputStream);

            var paragraphs = GetParagraphs(document, model.RangeType);
            if(model.IsApplyAction) {
                var shading = ColorTranslator.FromHtml(model.ShadingHex ?? string.Empty);
                FormatParagraphs(paragraphs, shading, model.ParagraphAlignment, model.LineSpacingMultiplier,
                    model.LeftIndent * DocumentUnitsPerInch, model.WidowOrphanControl, model.KeepWithNext, model.KeepLinesTogether);
                if(model.ApplyParagraphBorders) {
                    var borderColor = ColorTranslator.FromHtml(model.BorderColorHex ?? string.Empty);
                    FormatParagraphBorders(paragraphs, model.BorderLineWidth * DocumentUnitsPerPoint, model.BorderLineStyle, borderColor);
                }
            }
            else
                ResetParagraphFormatting(paragraphs, model.ParagraphResetValues ?? Array.Empty<ParagraphFormattingAttribute>());

            var outputStream = new MemoryStream();
            if(model.OutputFormat == "pdf")
                document.ExportToPdf(outputStream);
            else
                document.Save(outputStream, NewWordFormatHelper.GetDocumentFormat(model.OutputFormat));
            outputStream.Position = 0;
            return outputStream;
        }

        #region Code
        List<Paragraph> GetParagraphs(WordDocument document, ParagraphFormattingRangeType rangeType) {
            var result = new List<Paragraph>();
            if(rangeType == ParagraphFormattingRangeType.FirstParagraph) {
                var firstParagraph = FindFirstParagraph(document.Body);
                if(firstParagraph != null)
                    result.Add(firstParagraph);
            }
            else if(rangeType == ParagraphFormattingRangeType.FirstSection)
                CollectParagraphs(GetFirstSectionBlocks(document), result);
            else
                CollectParagraphs(document.Body, result);

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

        void CollectParagraphs(IEnumerable<Block> blocks, List<Paragraph> result) {
            foreach(var block in blocks) {
                if(block is Paragraph paragraph)
                    result.Add(paragraph);
                else if(block is Table table) {
                    foreach(var row in table.Rows) {
                        foreach(var cell in row.Cells)
                            CollectParagraphs(cell.Content, result);
                    }
                }
            }
        }

        void FormatParagraphs(IEnumerable<Paragraph> paragraphs, Color shading, ParagraphAlignment alignment,
                              float lineSpacingMultiplier, float leftIndent, bool widowOrphanControl, bool keepWithNext, bool keepLinesTogether) {
            foreach(var paragraph in paragraphs) {
                var properties = paragraph.Properties;
                properties.BackColor = new OfficeColor(shading);
                properties.Alignment = alignment;
                properties.LineSpacing = ParagraphLineSpacing.Multiple(lineSpacingMultiplier);
                properties.LeftIndent = leftIndent;
                properties.WidowOrphanControl = widowOrphanControl;
                properties.KeepWithNext = keepWithNext;
                properties.KeepLinesTogether = keepLinesTogether;
            }
        }

        void ResetParagraphFormatting(IEnumerable<Paragraph> paragraphs, IEnumerable<ParagraphFormattingAttribute> attributes) {
            foreach(var paragraph in paragraphs) {
                var properties = paragraph.Properties;
                foreach(var attribute in attributes) {
                    switch(attribute) {
                        case ParagraphFormattingAttribute.Alignment:
                            properties.Alignment = null;
                            break;
                        case ParagraphFormattingAttribute.LineSpacing:
                            properties.LineSpacing = null;
                            break;
                        case ParagraphFormattingAttribute.SpacingBefore:
                            properties.SpacingBefore = null;
                            break;
                        case ParagraphFormattingAttribute.SpacingAfter:
                            properties.SpacingAfter = null;
                            break;
                        case ParagraphFormattingAttribute.LeftIndent:
                            properties.LeftIndent = null;
                            break;
                        case ParagraphFormattingAttribute.Shading:
                            properties.BackColor = null;
                            break;
                        case ParagraphFormattingAttribute.WidowOrphanControl:
                            properties.WidowOrphanControl = null;
                            break;
                        case ParagraphFormattingAttribute.KeepWithNext:
                            properties.KeepWithNext = null;
                            break;
                        case ParagraphFormattingAttribute.KeepLinesTogether:
                            properties.KeepLinesTogether = null;
                            break;
                    }
                }
            }
        }

        void FormatParagraphBorders(IEnumerable<Paragraph> paragraphs, float lineWidth, BorderLineStyle lineStyle, Color borderColor) {
            foreach(var paragraph in paragraphs) {
                var borders = paragraph.Properties.Borders;
                SetParagraphBorder(borders.HorizontalBorder, lineWidth, lineStyle, borderColor);
                SetParagraphBorder(borders.BottomBorder, lineWidth, lineStyle, borderColor);
                SetParagraphBorder(borders.TopBorder, lineWidth, lineStyle, borderColor);
                SetParagraphBorder(borders.LeftBorder, lineWidth, lineStyle, borderColor);
                SetParagraphBorder(borders.RightBorder, lineWidth, lineStyle, borderColor);
            }
        }

        void SetParagraphBorder(ParagraphBorder border, float lineWidth, BorderLineStyle lineStyle, Color borderColor) {
            border.LineWidth = lineWidth;
            border.LineStyle = lineStyle;
            border.LineColor = new OfficeColor(borderColor);
        }
        #endregion Code
    }
}