using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "HeaderFooter_template.docx");
    const string OutputFileName = "Modified_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ProcessHeadersFootersDocument(inputStream, OutputFormat);
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

    static Stream ProcessHeadersFootersDocument(Stream inputStream, DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(inputStream);
        var document = wordProcessor.Document;
        document.Sections[0].DifferentFirstPage = true;
        AddHeader(wordProcessor, HeaderFooterType.First, "Cover Page", ParagraphAlignment.Center);
        AddFooter(wordProcessor, HeaderFooterType.First, "Internal Use Only", ParagraphAlignment.Center, true);
        AddHeader(wordProcessor, HeaderFooterType.Primary, "Document Title", ParagraphAlignment.Center);
        AddFooter(wordProcessor, HeaderFooterType.Primary, "Confidential", ParagraphAlignment.Center, true);

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static void AddHeader(RichEditDocumentServer wordProcessor, HeaderFooterType headerFooterType, string text, ParagraphAlignment alignment) {
        foreach(Section section in wordProcessor.Document.Sections) {
            var header = section.BeginUpdateHeader(headerFooterType);
            header.InsertText(header.Paragraphs[0].Range.Start, text);
            var properties = header.BeginUpdateParagraphs(header.Range);
            properties.Alignment = alignment;
            header.EndUpdateParagraphs(properties);
            section.EndUpdateHeader(header);
        }
    }

    static void AddFooter(RichEditDocumentServer wordProcessor, HeaderFooterType headerFooterType, string text, ParagraphAlignment alignment, bool addPageNumbering) {
        foreach(Section section in wordProcessor.Document.Sections) {
            var footer = section.BeginUpdateFooter(headerFooterType);
            footer.InsertText(footer.Paragraphs[0].Range.Start, text);
            if(addPageNumbering) {
                footer.Paragraphs.Append();
                footer.AppendText("Page ");
                footer.Fields.Create(footer.Range.End, "PAGE");
                footer.AppendText(" of ");
                footer.Fields.Create(footer.Range.End, "NUMPAGES");
            }
            var properties = footer.BeginUpdateParagraphs(footer.Range);
            properties.Alignment = alignment;
            footer.EndUpdateParagraphs(properties);
            section.EndUpdateFooter(footer);
        }
    }
}
