using DevExpress.Docs.Presentation;
using DevExpress.Docs.Office;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.pptx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Encrypted_Presentation.pptx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Pptx;
    const string Password = "123";
    const EncryptionType EncryptionType = DevExpress.Docs.Office.EncryptionType.Strong;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = EncryptDocumentWithPassword(inputStream, OutputFormat, Password, EncryptionType);
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

    static Stream EncryptDocumentWithPassword(Stream inputStream, DocumentFormat outputFormat, string password, EncryptionType encryptionType) {
        using var presentation = new Presentation(inputStream);

        var options = new EncryptionOptions(password, encryptionType);
        presentation.Encrypt(options);

        var outputStream = new MemoryStream();
        presentation.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }
}
