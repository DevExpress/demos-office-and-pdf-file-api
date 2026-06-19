using System.Text;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputFileName = "Document_Statistics.txt";
    const bool IncludeHeaderFooter = true;
    const bool IncludeTextBoxes = true;
    const bool IncludeFootnotesEndnotes = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            var statisticsText = GenerateDocumentStatisticsText(inputStream, IncludeHeaderFooter, IncludeTextBoxes, IncludeFootnotesEndnotes);
            File.WriteAllText(outputPath, statisticsText, Encoding.UTF8);
            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static string GenerateDocumentStatisticsText(Stream inputStream, bool includeHeaderFooter, bool includeTextBoxes, bool includeFootnotesEndnotes) {
        var output = new StringBuilder();
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(inputStream);
        var document = wordProcessor.Document;

        output.AppendLine("BASIC DOCUMENT STATISTICS:");
        output.AppendLine(new string('-', 35));

        var pageCount = wordProcessor.DocumentLayout.GetPageCount();
        var sectionCount = document.Sections.Count;
        var paragraphCount = 0;
        var totalWords = 0;
        var totalChars = 0;
        var totalCharsNoSpaces = 0;
        var totalTables = 0;
        var totalImages = 0;
        var totalHyperlinks = 0;
        var totalBookmarks = 0;

        ForEachSubDocument(document, subDocument => {
            var subDocumentType = subDocument.GetSubDocumentType();
            if(subDocumentType == SubDocumentType.Main ||
               includeHeaderFooter && (subDocumentType == SubDocumentType.Header || subDocumentType == SubDocumentType.Footer) ||
               includeFootnotesEndnotes && (subDocumentType == SubDocumentType.FootNote || subDocumentType == SubDocumentType.EndNote) ||
               includeTextBoxes && subDocumentType == SubDocumentType.TextBox) {
                paragraphCount += subDocument.Paragraphs.Count;
                var text = subDocument.GetText(subDocument.Range);
                totalWords += CountWords(text);
                totalChars += text.Length;
                totalCharsNoSpaces += text.Replace(" ", string.Empty).Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\t", string.Empty).Length;
                totalTables += subDocument.Tables.Count;
                totalImages += subDocument.Shapes.Count;
                totalHyperlinks += subDocument.Hyperlinks.Count;
                totalBookmarks += subDocument.Bookmarks.Count;
            }
        });

        output.AppendLine($"Pages: {pageCount}");
        output.AppendLine($"Sections: {sectionCount}");
        output.AppendLine($"Paragraphs: {paragraphCount}");
        output.AppendLine();
        output.AppendLine("TEXT CONTENT ANALYSIS:");
        output.AppendLine(new string('-', 35));
        output.AppendLine($"Words: {totalWords:N0}");
        output.AppendLine($"Characters (with spaces): {totalChars:N0}");
        output.AppendLine($"Characters (no spaces): {totalCharsNoSpaces:N0}");
        output.AppendLine();
        output.AppendLine("DOCUMENT ELEMENTS:");
        output.AppendLine(new string('-', 35));
        output.AppendLine($"Tables: {totalTables:N0}");
        output.AppendLine($"Images/Shapes: {totalImages:N0}");
        output.AppendLine($"Hyperlinks: {totalHyperlinks:N0}");
        output.AppendLine($"Bookmarks: {totalBookmarks:N0}");
        output.AppendLine($"Comments: {document.Comments.Count:N0}");
        output.AppendLine($"Footnotes: {document.Footnotes.Count:N0}");
        output.AppendLine($"Endnotes: {document.Endnotes.Count:N0}");
        output.AppendLine();
        return output.ToString();
    }

    static int CountWords(string text) {
        if(string.IsNullOrWhiteSpace(text))
            return 0;
        return text.Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
    }

    delegate void SubDocumentProcessor(SubDocument subDocument);

    static void ForEachSubDocument(Document document, SubDocumentProcessor processor) {
        processor(document);
        ProcessNotes(document.Footnotes, processor);
        ProcessNotes(document.Endnotes, processor);
        ProcessShapes(document.Shapes, processor);
        ProcessComments(document.Comments, processor);
        foreach(Section section in document.Sections) {
            ProcessSection(section, HeaderFooterType.First, processor);
            ProcessSection(section, HeaderFooterType.Odd, processor);
            ProcessSection(section, HeaderFooterType.Even, processor);
        }
    }

    static void ProcessSection(Section section, HeaderFooterType headerFooterType, SubDocumentProcessor processor) {
        if(section.HasHeader(headerFooterType)) {
            var header = section.BeginUpdateHeader(headerFooterType);
            processor(header);
            ProcessShapes(header.Shapes, processor);
            section.EndUpdateHeader(header);
        }
        if(section.HasFooter(headerFooterType)) {
            var footer = section.BeginUpdateFooter(headerFooterType);
            processor(footer);
            ProcessShapes(footer.Shapes, processor);
            section.EndUpdateFooter(footer);
        }
    }

    static void ProcessShapes(ShapeCollection shapes, SubDocumentProcessor processor) {
        foreach(Shape shape in shapes) {
            if(shape.ShapeFormat != null)
                processor(shape.ShapeFormat.TextBox.Document);
        }
    }

    static void ProcessComments(CommentCollection comments, SubDocumentProcessor processor) {
        foreach(Comment comment in comments) {
            var commentDocument = comment.BeginUpdate();
            processor(commentDocument);
            comment.EndUpdate(commentDocument);
        }
    }

    static void ProcessNotes(NoteCollection notes, SubDocumentProcessor processor) {
        foreach(Note note in notes) {
            var noteDocument = note.BeginUpdate();
            processor(noteDocument);
            note.EndUpdate(noteDocument);
        }
    }
}
