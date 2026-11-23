using DevExpress.Docs.Office;
using DevExpress.Docs.Word;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Encrypted_Document.docx";
    const string Password = "demo-password";
    const EncryptionType DocumentEncryptionType = EncryptionType.Strong;
    const DocumentFormat OutputFormat = DocumentFormat.Docx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = EncryptDocumentWithPassword(inputStream, OutputFormat, Password, DocumentEncryptionType);
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
        // Load document from stream
        using var document = new WordDocument(inputStream);

        // Create encryption settings
        document.Encrypt(new EncryptionOptions(password, encryptionType));

        // Prepare output stream
        var outputStream = new MemoryStream();

        // Save the encrypted document
        document.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }
}