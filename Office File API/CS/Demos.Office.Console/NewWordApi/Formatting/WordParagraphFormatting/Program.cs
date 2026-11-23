using DevExpress.Docs;
using DevExpress.Docs.Word;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    static readonly string OutputFormat = "docx";
    static readonly string OutputFileName = "Formatted_Document." + OutputFormat;
    static readonly ParagraphFormattingActionType ActionType = ParagraphFormattingActionType.ApplyFormatting;
    const ParagraphFormattingRangeType RangeType = ParagraphFormattingRangeType.FirstParagraph;
    const ParagraphAlignment Alignment = ParagraphAlignment.Left;
    const float LineSpacingMultiplier = 1.15f;
    const float LeftIndentInches = 0f;
    const string ShadingHex = "#FFFFFF";
    const bool WidowOrphanControl = true;
    const bool KeepWithNext = false;
    const bool KeepLinesTogether = false;
    static readonly bool ApplyParagraphBorders = false;
    const float BorderLineWidthPoints = 2f;
    const BorderLineStyle ParagraphBorderLineStyle = BorderLineStyle.Single;
    const string BorderColorHex = "#4682B4";
    static readonly ParagraphFormattingAttribute[] ResetAttributes = Enum.GetValues<ParagraphFormattingAttribute>();

    const float DocumentUnitsPerInch = 300f;
    const float DocumentUnitsPerPoint = DocumentUnitsPerInch / 72f;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = FormatDocument(inputStream);
            using var fileStream = File.Create(outputPath);
            outputStream.CopyTo(fileStream);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static Stream FormatDocument(Stream inputStream) {
        using var document = new WordDocument(inputStream);

        var paragraphs = GetParagraphs(document, RangeType);
        if(ActionType == ParagraphFormattingActionType.ApplyFormatting) {
            var shading = ColorTranslator.FromHtml(ShadingHex);
            FormatParagraphs(paragraphs, shading, Alignment, LineSpacingMultiplier,
                LeftIndentInches * DocumentUnitsPerInch, WidowOrphanControl, KeepWithNext, KeepLinesTogether);
            if(ApplyParagraphBorders) {
                var borderColor = ColorTranslator.FromHtml(BorderColorHex);
                FormatParagraphBorders(paragraphs, BorderLineWidthPoints * DocumentUnitsPerPoint, ParagraphBorderLineStyle, borderColor);
            }
        }
        else {
            ResetParagraphFormatting(paragraphs, ResetAttributes);
        }

        var outputStream = new MemoryStream();
        if(OutputFormat == "pdf") {
            document.ExportToPdf(outputStream);
        }
        else {
            document.Save(outputStream, GetDocumentFormat(OutputFormat));
        }
        outputStream.Position = 0;
        return outputStream;
    }

    static List<Paragraph> GetParagraphs(WordDocument document, ParagraphFormattingRangeType rangeType) {
        var result = new List<Paragraph>();
        if(rangeType == ParagraphFormattingRangeType.FirstParagraph) {
            var firstParagraph = FindFirstParagraph(document.Body);
            if(firstParagraph != null) {
                result.Add(firstParagraph);
            }
        }
        else if(rangeType == ParagraphFormattingRangeType.FirstSection) {
            CollectParagraphs(GetFirstSectionBlocks(document), result);
        }
        else {
            CollectParagraphs(document.Body, result);
        }

        return result;
    }

    static List<Block> GetFirstSectionBlocks(WordDocument document) {
        var result = new List<Block>();
        foreach(var block in document.Body) {
            result.Add(block);
            if(block is Paragraph paragraph && paragraph.Mark.Type != ParagraphMarkType.ParagraphMark) {
                break;
            }
        }

        return result;
    }

    static Paragraph? FindFirstParagraph(IEnumerable<Block> blocks) {
        foreach(var block in blocks) {
            if(block is Paragraph paragraph) {
                return paragraph;
            }
            if(block is Table table) {
                foreach(var row in table.Rows) {
                    foreach(var cell in row.Cells) {
                        var cellParagraph = FindFirstParagraph(cell.Content);
                        if(cellParagraph != null) {
                            return cellParagraph;
                        }
                    }
                }
            }
        }

        return null;
    }

    static void CollectParagraphs(IEnumerable<Block> blocks, List<Paragraph> result) {
        foreach(var block in blocks) {
            if(block is Paragraph paragraph) {
                result.Add(paragraph);
            }
            else if(block is Table table) {
                foreach(var row in table.Rows) {
                    foreach(var cell in row.Cells) {
                        CollectParagraphs(cell.Content, result);
                    }
                }
            }
        }
    }

    static void FormatParagraphs(IEnumerable<Paragraph> paragraphs, Color shading, ParagraphAlignment alignment,
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

    static void ResetParagraphFormatting(IEnumerable<Paragraph> paragraphs, IEnumerable<ParagraphFormattingAttribute> attributes) {
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

    static void FormatParagraphBorders(IEnumerable<Paragraph> paragraphs, float lineWidth, BorderLineStyle lineStyle, Color borderColor) {
        foreach(var paragraph in paragraphs) {
            var borders = paragraph.Properties.Borders;
            SetParagraphBorder(borders.HorizontalBorder, lineWidth, lineStyle, borderColor);
            SetParagraphBorder(borders.BottomBorder, lineWidth, lineStyle, borderColor);
            SetParagraphBorder(borders.TopBorder, lineWidth, lineStyle, borderColor);
            SetParagraphBorder(borders.LeftBorder, lineWidth, lineStyle, borderColor);
            SetParagraphBorder(borders.RightBorder, lineWidth, lineStyle, borderColor);
        }
    }

    static void SetParagraphBorder(ParagraphBorder border, float lineWidth, BorderLineStyle lineStyle, Color borderColor) {
        border.LineWidth = lineWidth;
        border.LineStyle = lineStyle;
        border.LineColor = new OfficeColor(borderColor);
    }

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "doc" => DocumentFormat.Doc,
            "docx" => DocumentFormat.Docx,
            "dotx" => DocumentFormat.Dotx,
            "rtf" => DocumentFormat.Rtf,
            _ => DocumentFormat.Docx
        };
    }

    enum ParagraphFormattingActionType {
        ApplyFormatting,
        ClearFormatting
    }

    enum ParagraphFormattingRangeType {
        FirstParagraph,
        FirstSection,
        WholeDocument
    }

    enum ParagraphFormattingAttribute {
        Alignment,
        LineSpacing,
        SpacingBefore,
        SpacingAfter,
        LeftIndent,
        Shading,
        WidowOrphanControl,
        KeepWithNext,
        KeepLinesTogether
    }
}