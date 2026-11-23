using DevExpress.Docs.Word;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "ConvertedDocument.docx";
    const string OutputFormat = "docx";
    const bool EmbedFonts = false;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ConvertDocument(inputStream, EmbedFonts, GetDocumentFormat(OutputFormat));
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

    static Stream ConvertDocument(Stream inputStream, bool embedFonts, DocumentFormat outputFormat) {
        // Load the input document
        using var document = new WordDocument(inputStream);

        document.EmbedFonts = embedFonts;

        var saveOptions = new SaveOptions() {
            DocumentFormat = outputFormat
        };
        // Configure HTML export settings for image handling
        saveOptions.HtmlOptions.EmbedImages = true;
        saveOptions.MhtOptions.EmbedImages = true;

        // Create output stream
        var outputStream = new MemoryStream();

        // Save document in target format
        document.Save(outputStream, saveOptions);

        outputStream.Position = 0;
        return outputStream;
    }

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "doc" => DocumentFormat.Doc,
            "docx" => DocumentFormat.Docx,
            "docm" => DocumentFormat.Docm,
            "dotx" => DocumentFormat.Dotx,
            "dotm" => DocumentFormat.Dotm,
            "rtf" => DocumentFormat.Rtf,
            "html" => DocumentFormat.Html,
            "mht" => DocumentFormat.Mht,
            "txt" => DocumentFormat.PlainText,
            "odt" => DocumentFormat.Odt,
            _ => DocumentFormat.Docx
        };
    }
}