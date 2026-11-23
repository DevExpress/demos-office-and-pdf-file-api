using DevExpress.Docs.Word;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Protected_Document.docx";
    const string Password = "demo-password";
    const DocumentProtectionType ProtectionType = DocumentProtectionType.ReadOnly;
    const DocumentFormat OutputFormat = DocumentFormat.Docx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ProtectDocumentWithPassword(inputStream, OutputFormat, Password, ProtectionType);
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

    static Stream ProtectDocumentWithPassword(Stream inputStream, DocumentFormat outputFormat, string password, DocumentProtectionType protectionType) {
        using var document = new WordDocument(inputStream);

        // Unprotect document if it's already protected
        if(document.ProtectionType != DocumentProtectionType.None) {
            document.RemoveProtection();
        }

        // Apply document protection with password
        var protectionOptions = new DocumentProtectionOptions(password) {
            ProtectionType = protectionType
        };
        document.Protect(protectionOptions);

        // Save protected document
        var outputStream = new MemoryStream();

        document.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }
}