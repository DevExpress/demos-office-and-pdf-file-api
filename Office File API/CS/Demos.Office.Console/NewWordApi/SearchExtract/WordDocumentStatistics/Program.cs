using System.Text;
using DevExpress.Docs.Word;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Document_Statistics.txt";
    const bool IncludeHeaderFooter = true;
    const bool IncludeTextBoxes = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            var statisticsText = GenerateDocumentStatisticsText(inputStream, IncludeHeaderFooter, IncludeTextBoxes);
            File.WriteAllText(outputPath, statisticsText, Encoding.UTF8);
            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static string GenerateDocumentStatisticsText(Stream inputStream, bool includeHeaderFooter, bool includeTextBoxes) {
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

    static void CollectBlocks(IEnumerable<Block> blocks, List<Paragraph> paragraphs, List<Table> tables) {
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

    static void CollectInlineElements(IEnumerable<DocumentElement> elements, List<Inline> result) {
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

    static void CollectTextBoxShapes(IEnumerable<ShapeBase> shapes, List<Shape> result) {
        foreach(var shape in shapes) {
            if(shape is Shape textBoxShape)
                result.Add(textBoxShape);
            else if(shape is ContainerShapeBase containerShape)
                CollectTextBoxShapes(containerShape.Shapes, result);
        }
    }

    static int CountBookmarks(WordDocument document, bool includeTextBoxes) {
        var result = document.Bookmarks.Count(bookmark =>
            bookmark.DocumentFragment == document.Body
            || (includeTextBoxes && bookmark.DocumentFragment is ShapeTextBoxContent));
        return result;
    }

    static int CountWords(string text) {
        if(string.IsNullOrWhiteSpace(text))
            return 0;

        return text.Split(new char[] { ' ', '\r', '\n', '\t', '\v' }, StringSplitOptions.RemoveEmptyEntries).Length;
    }
}