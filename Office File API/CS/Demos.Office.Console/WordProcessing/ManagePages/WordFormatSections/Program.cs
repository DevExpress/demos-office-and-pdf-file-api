using DevExpress.Drawing.Printing;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputFileName = "FormattedDocument.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = FormatDocumentSections(inputStream, OutputFormat);
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

    static Stream FormatDocumentSections(Stream inputStream, DocumentFormat outputFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(inputStream);
        var document = wordProcessor.Document;

        foreach(Section section in document.Sections)
            ApplySectionFormatting(document, section);

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static void ApplySectionFormatting(Document document, Section section) {
        SetPaperSizeAndOrientation(document, section, DXPaperKind.Letter, 8.5f, 11f, false);
        SetMargins(document, section, 1f, 1f, 1f, 1f, false);
        SetHeaderFooterOffsets(document, section, 0.5f, 0.5f);
        SetGutterSettings(document, section, false, GutterPosition.Left, 0f);
    }

    static void SetPaperSizeAndOrientation(Document document, Section section, DXPaperKind paperKind, float customPageWidth, float customPageHeight, bool isLandscape) {
        document.Unit = DevExpress.Office.DocumentUnit.Inch;
        section.Page.PaperKind = paperKind;
        if(paperKind == DXPaperKind.Custom) {
            section.Page.Width = customPageWidth;
            section.Page.Height = customPageHeight;
        }
        section.Page.Landscape = isLandscape;
    }

    static void SetMargins(Document document, Section section, float leftMargin, float rightMargin, float topMargin, float bottomMargin, bool mirroredMargins) {
        document.Unit = DevExpress.Office.DocumentUnit.Inch;
        section.Margins.Top = topMargin;
        section.Margins.Bottom = bottomMargin;
        section.Margins.Left = leftMargin;
        section.Margins.Right = rightMargin;
        document.MarginType = mirroredMargins ? MarginType.Mirrored : MarginType.Normal;
    }

    static void SetHeaderFooterOffsets(Document document, Section section, float headerOffset, float footerOffset) {
        document.Unit = DevExpress.Office.DocumentUnit.Inch;
        section.Margins.HeaderOffset = headerOffset;
        section.Margins.FooterOffset = footerOffset;
    }

    static void SetGutterSettings(Document document, Section section, bool isGutterAtTop, GutterPosition gutterPosition, float gutter) {
        document.Unit = DevExpress.Office.DocumentUnit.Inch;
        section.Margins.GutterPosition = gutterPosition;
        section.Margins.Gutter = gutter;
        document.GutterAtTop = isGutterAtTop;
    }
}
