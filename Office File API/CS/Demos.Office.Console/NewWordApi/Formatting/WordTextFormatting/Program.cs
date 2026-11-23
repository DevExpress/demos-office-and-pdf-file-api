using DevExpress.Docs;
using DevExpress.Docs.Office;
using DevExpress.Docs.Word;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    static readonly string OutputFormat = "docx";
    static readonly string OutputFileName = "Formatted_Document." + OutputFormat;
    static readonly TextFormattingActionType ActionType = TextFormattingActionType.ApplyFormatting;
    const TextFormattingRangeType RangeType = TextFormattingRangeType.FirstParagraph;
    const string FontName = "Segoe UI";
    const float FontSize = 16f;
    static readonly TextFontStyle FontStyle = TextFontStyle.Regular;
    const string ForeColorHex = "#000000";
    const string BackColorHex = "#FFFFFF";
    const TextUnderlineType Underline = TextUnderlineType.None;
    static readonly TextFormattingAttribute[] ResetAttributes = Enum.GetValues<TextFormattingAttribute>();

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

        var textFragments = GetTextFragments(document, RangeType);
        if(ActionType == TextFormattingActionType.ApplyFormatting) {
            var foreColor = ColorTranslator.FromHtml(ForeColorHex);
            var backColor = ColorTranslator.FromHtml(BackColorHex);
            var isBold = FontStyle is TextFontStyle.Bold or TextFontStyle.BoldItalic;
            var isItalic = FontStyle is TextFontStyle.Italic or TextFontStyle.BoldItalic;
            FormatText(textFragments, FontName, FontSize, foreColor, backColor, Underline, isBold, isItalic);
        }
        else {
            ResetCharacterFormatting(textFragments, ResetAttributes);
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

    static List<TextFragment> GetTextFragments(WordDocument document, TextFormattingRangeType rangeType) {
        var result = new List<TextFragment>();
        if(rangeType == TextFormattingRangeType.FirstParagraph) {
            var firstParagraph = FindFirstParagraph(document.Body);
            if(firstParagraph != null) {
                CollectTextFragments(firstParagraph.InlineElements, result);
            }
        }
        else if(rangeType == TextFormattingRangeType.FirstSection) {
            CollectTextFragments(GetFirstSectionBlocks(document), result);
        }
        else {
            CollectTextFragments(document.Body, result);
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

    static void CollectTextFragments(IEnumerable<DocumentElement> elements, List<TextFragment> result) {
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
                        foreach(var cell in row.Cells) {
                            CollectTextFragments(cell.Content, result);
                        }
                    }

                    break;
                }
            }
        }
    }

    static void FormatText(IEnumerable<TextFragment> textFragments, string fontName, float fontSize,
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

    static void ResetCharacterFormatting(IEnumerable<TextFragment> textFragments, IEnumerable<TextFormattingAttribute> attributes) {
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

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "doc" => DocumentFormat.Doc,
            "docx" => DocumentFormat.Docx,
            "dotx" => DocumentFormat.Dotx,
            "rtf" => DocumentFormat.Rtf,
            _ => DocumentFormat.Docx
        };
    }

    enum TextFormattingActionType {
        ApplyFormatting,
        ClearFormatting
    }

    enum TextFormattingRangeType {
        FirstParagraph,
        FirstSection,
        WholeDocument
    }

    enum TextFontStyle {
        Regular,
        Italic,
        Bold,
        BoldItalic
    }

    enum TextFormattingAttribute {
        FontName,
        FontSize,
        ForeColor,
        Italic,
        Bold,
        BackColor,
        Underline
    }
}