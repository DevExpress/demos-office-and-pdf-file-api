using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "FootnoteEndnote_template.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "FootnotesEndnotes_Result.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const bool AddNewFootnote = true;
    const bool AddNewEndnote = true;
    const string NoteSeparator = " - ";
    const NoteRestartType NoteRestartType = DevExpress.XtraRichEdit.API.Native.NoteRestartType.Continuous;
    const NumberingFormat NumberingFormat = DevExpress.XtraRichEdit.API.Native.NumberingFormat.Decimal;
    const int StartNumber = 1;
    const EndnotePosition EndnotePosition = DevExpress.XtraRichEdit.API.Native.EndnotePosition.EndOfDocument;
    const FootnotePosition FootnotePosition = DevExpress.XtraRichEdit.API.Native.FootnotePosition.BottomOfPage;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ProcessFootnotesEndnotesDocument(inputStream, OutputFormat);
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

    static Stream ProcessFootnotesEndnotesDocument(Stream templateStream, DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(templateStream);
        var document = wordProcessor.Document;

        if(AddNewFootnote)
            AddFootnote(document, document.Bookmarks["footnote"].Range.Start, "This is a new footnote");
        if(AddNewEndnote)
            AddEndnote(document, document.Bookmarks["endnote"].Range.Start, "This is a new endnote");

        ChangeNoteAppearance(document);
        ChangeNotePosition(document);

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static void AddFootnote(Document document, DocumentPosition position, string text) {
        var footnote = document.Footnotes.Insert(position);
        var footnoteDocument = footnote.BeginUpdate();
        var range = footnoteDocument.AppendText(text);
        var properties = footnoteDocument.BeginUpdateCharacters(range);
        properties.Bold = true;
        properties.ForeColor = Color.Red;
        footnoteDocument.EndUpdateCharacters(properties);
        footnote.EndUpdate(footnoteDocument);
    }

    static void AddEndnote(Document document, DocumentPosition position, string text) {
        var endnote = document.Endnotes.Insert(position);
        var endnoteDocument = endnote.BeginUpdate();
        var range = endnoteDocument.AppendText(text);
        var properties = endnoteDocument.BeginUpdateCharacters(range);
        properties.Bold = true;
        properties.ForeColor = Color.Red;
        endnoteDocument.EndUpdateCharacters(properties);
        endnote.EndUpdate(endnoteDocument);
    }

    static void ChangeNoteAppearance(Document document) {
        var separatorDocument = document.Footnotes.BeginUpdateSeparator(NoteSeparatorType.Separator);
        separatorDocument.Delete(separatorDocument.Range);
        separatorDocument.AppendText(NoteSeparator);
        document.Footnotes.EndUpdateSeparator(separatorDocument);

        foreach(Section section in document.Sections) {
            var options = section.FootnoteOptions;
            options.RestartType = NoteRestartType;
            options.NumberingFormat = NumberingFormat;
            options.StartNumber = StartNumber;
        }
    }

    static void ChangeNotePosition(Document document) {
        document.EndnotePosition = EndnotePosition;
        foreach(Section section in document.Sections)
            section.FootnoteOptions.Position = FootnotePosition;
    }
}
