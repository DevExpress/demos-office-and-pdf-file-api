using DevExpress.Docs.Presentation;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.pptx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Converted_Presentation.pptx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Pptx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ConvertToFormat(inputStream, OutputFormat);
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

    static Stream ConvertToFormat(Stream inputStream, DocumentFormat targetFormat) {
        using var presentation = new Presentation(inputStream);

        var outputStream = new MemoryStream();
        presentation.Save(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
