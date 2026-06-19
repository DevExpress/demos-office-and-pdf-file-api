using System.Text;
using DevExpress.Office;
using DevExpress.XtraRichEdit;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputFormat = "txt";
    const InspectOperationType Operation = InspectOperationType.Inspect;
    const string InspectOutputFileName = "Inspection_Log.txt";
    const string SanitizeOutputFileName = "Sanitized_Document.docx";
    const bool RemoveHiddenText = true;
    const bool RemoveInvisibleText = true;
    const bool RemoveMetadata = true;
    const bool RemoveComments = true;
    const bool RemoveCustomXmlParts = true;
    const bool RemoveActiveXContent = true;
    const bool RemoveMacros = true;
    const bool RemoveOleObjects = true;
    const bool RemoveTrackChanges = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputFileName = Operation == InspectOperationType.Inspect ? InspectOutputFileName : SanitizeOutputFileName;
            var outputPath = Path.Combine(AppContext.BaseDirectory, outputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = Operation == InspectOperationType.Inspect
                ? CreateInspectionLog(inputStream)
                : SanitizeDocument(inputStream, GetDocumentFormat(OutputFormat), new WordProcessingSanitizeOptions {
                    HiddenText = RemoveHiddenText ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    InvisibleText = RemoveInvisibleText ? InvisibleContentSanitizeMode.Remove : InvisibleContentSanitizeMode.Ignore,
                    TrackedChanges = RemoveTrackChanges ? TrackedChangesSanitizeMode.Reject : TrackedChangesSanitizeMode.Ignore,
                    RemoveMacros = RemoveMacros,
                    RemoveActiveXContent = RemoveActiveXContent,
                    RemoveOleObjects = RemoveOleObjects,
                    RemoveComments = RemoveComments,
                    RemoveCustomXmlParts = RemoveCustomXmlParts,
                    Metadata = RemoveMetadata ? MetadataRemovalScope.All : MetadataRemovalScope.None,
                });
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

    static MemoryStream CreateInspectionLog(Stream inputStream) {
        var inspectionResult = InspectDocument(inputStream, WordProcessingInspectOptions.All);
        var outputStream = new MemoryStream();
        using var writer = new StreamWriter(outputStream, Encoding.UTF8, leaveOpen: true);
        writer.Write(inspectionResult);
        writer.Flush();
        outputStream.Position = 0;
        return outputStream;
    }

    static string InspectDocument(Stream inputStream, WordProcessingInspectOptions inspectOptions) {
        using var richEditDocumentServer = new RichEditDocumentServer();
        richEditDocumentServer.LoadDocument(inputStream);

        StringBuilder output = new StringBuilder();
        var inspectResult = richEditDocumentServer.Inspect(inspectOptions);
        if(inspectResult.ContentTypes.Count == 0)
            output.AppendLine("No issues found.");
        else {
            output.AppendLine("Document contains the following hidden/sensitive content:");
            output.AppendLine(string.Join(Environment.NewLine, inspectResult.ContentTypes));
        }
        return output.ToString();
    }

    static MemoryStream SanitizeDocument(Stream inputStream, DocumentFormat outputFormat, WordProcessingSanitizeOptions? sanitizeOptions = null) {
        using var richEditDocumentServer = new RichEditDocumentServer();
        richEditDocumentServer.LoadDocument(inputStream);

        if(sanitizeOptions == null) {
            var inspectResult = richEditDocumentServer.Inspect();
            sanitizeOptions = inspectResult.CreateSanitizeOptions();
        }

        richEditDocumentServer.Sanitize(sanitizeOptions);

        var outputStream = new MemoryStream();
        richEditDocumentServer.SaveDocument(outputStream, outputFormat);
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
            "txt" => DocumentFormat.PlainText,
            "html" => DocumentFormat.Html,
            "mht" => DocumentFormat.Mht,
            "odt" => DocumentFormat.OpenDocument,
            _ => DocumentFormat.Docx
        };
    }

    enum InspectOperationType {
        Inspect,
        Sanitize
    }
}
