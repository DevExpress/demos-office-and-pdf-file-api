using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputFileName = "EncryptedDocument.pdf";
    const EncryptionAlgorithm Algorithm = EncryptionAlgorithm.AES256;
    const DocumentDataExtractionPermissions DataExtractionPermissions = DocumentDataExtractionPermissions.Allowed;
    const DocumentInteractivityPermissions InteractivityPermissions = DocumentInteractivityPermissions.Allowed;
    const DocumentModificationPermissions ModificationPermissions = DocumentModificationPermissions.Allowed;
    const DocumentPrintPermissions PrintPermissions = DocumentPrintPermissions.Allowed;
    const string OwnerPassword = "dxdemo-owner";
    const string UserPassword = "dxdemo-user";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

            using var documentStream = File.OpenRead(inputPath);
            using var outputStream = Encrypt(documentStream, Algorithm, DataExtractionPermissions, InteractivityPermissions, ModificationPermissions, PrintPermissions, OwnerPassword, UserPassword);
            SaveDocument(outputStream, outputPath);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static void SaveDocument(Stream stream, string outputPath) {
        stream.Position = 0;
        using var fileStream = File.Create(outputPath);
        stream.CopyTo(fileStream);
    }

    static Stream Encrypt(Stream documentStream, EncryptionAlgorithm algorithm, DocumentDataExtractionPermissions dataExtractionPermissions, DocumentInteractivityPermissions interactivityPermissions, DocumentModificationPermissions modificationPermissions, DocumentPrintPermissions printPermissions, string ownerPassword, string userPassword) {
        using var doc = new PdfDocument(documentStream);

        doc.Encrypt(new EncryptionOptions(ownerPassword, userPassword) {
            Algorithm = algorithm,
            DataExtractionPermissions = dataExtractionPermissions,
            InteractivityPermissions = interactivityPermissions,
            ModificationPermissions = modificationPermissions,
            PrintPermissions = printPermissions
        });

        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
