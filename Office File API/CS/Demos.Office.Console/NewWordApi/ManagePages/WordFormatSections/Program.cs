using DevExpress.Docs.Word;
using DevExpress.Drawing.Printing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "FormattedDocument.docx";
    const string OutputFormat = "docx";
    const float DocumentUnitsPerInch = 300f;
    const DXPaperKind PaperSize = DXPaperKind.Letter;
    const bool IsLandscape = false;
    const float CustomPaperWidth = 8.5f;
    const float CustomPaperHeight = 11.0f;
    const float TopMargin = 1.0f;
    const float BottomMargin = 1.0f;
    const float LeftMargin = 1.0f;
    const float RightMargin = 1.0f;
    const bool MirroredMargins = false;
    const float HeaderOffset = 0.5f;
    const float FooterOffset = 0.5f;
    const float GutterSize = 0f;
    const string DefaultGutterPosition = "Left";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
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

    static Stream FormatDocumentSections(Stream inputStream, string outputFormat) {
        using var document = new WordDocument(inputStream);

        foreach(var section in document.Sections)
            ApplySectionFormatting(document, section);

        var outputStream = new MemoryStream();
        document.Save(outputStream, GetDocumentFormat(outputFormat));
        outputStream.Position = 0;
        return outputStream;
    }

    static void ApplySectionFormatting(WordDocument document, Section section) {
        var orientation = IsLandscape ? PageOrientation.Landscape : PageOrientation.Portrait;
        SetPaperSizeAndOrientation(section, PaperSize,
            ToDocumentUnits(CustomPaperWidth), ToDocumentUnits(CustomPaperHeight), orientation);

        SetMargins(document, section, ToDocumentUnits(LeftMargin), ToDocumentUnits(RightMargin),
            ToDocumentUnits(TopMargin), ToDocumentUnits(BottomMargin), MirroredMargins);

        SetHeaderFooterOffsets(section, ToDocumentUnits(HeaderOffset), ToDocumentUnits(FooterOffset));

        SetGutterSettings(document, section, IsGutterAtTop(DefaultGutterPosition), GetGutterPosition(DefaultGutterPosition), ToDocumentUnits(GutterSize));
    }

    static float ToDocumentUnits(float inches) {
        return inches * DocumentUnitsPerInch;
    }

    static bool IsGutterAtTop(string gutterPosition) {
        return gutterPosition.ToLowerInvariant() == "top";
    }

    static GutterPosition GetGutterPosition(string gutterPosition) {
        return gutterPosition.ToLowerInvariant() == "right" ? GutterPosition.Right : GutterPosition.Left;
    }

    static void SetPaperSizeAndOrientation(Section section, DXPaperKind paperKind, float customPageWidth, float customPageHeight,
                                           PageOrientation orientation) {
        if(paperKind == DXPaperKind.Custom) {
            section.Page.Orientation = PageOrientation.Portrait;
            section.Page.Width = customPageWidth;
            section.Page.Height = customPageHeight;
        }
        else {
            section.Page.PaperKind = paperKind;
        }

        section.Page.Orientation = orientation;
    }

    static void SetMargins(WordDocument document, Section section, float leftMargin, float rightMargin, float topMargin, float bottomMargin, bool mirroredMargins) {
        section.Margins.Top = topMargin;
        section.Margins.Bottom = bottomMargin;
        section.Margins.Left = leftMargin;
        section.Margins.Right = rightMargin;
        document.MirrorMargins = mirroredMargins;
    }

    static void SetHeaderFooterOffsets(Section section, float headerOffset, float footerOffset) {
        section.Margins.HeaderOffset = headerOffset;
        section.Margins.FooterOffset = footerOffset;
    }

    static void SetGutterSettings(WordDocument document, Section section, bool isGutterAtTop, GutterPosition gutterPosition, float gutter) {
        section.Margins.GutterPosition = gutterPosition;
        section.Margins.Gutter = gutter;
        document.GutterAtTop = isGutterAtTop;
    }

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "doc" => DocumentFormat.Doc,
            "docx" => DocumentFormat.Docx,
            "docm" => DocumentFormat.Docm,
            "dotx" => DocumentFormat.Dotx,
            "dotm" => DocumentFormat.Dotm,
            "rtf" => DocumentFormat.Rtf,
            _ => DocumentFormat.Docx
        };
    }
}