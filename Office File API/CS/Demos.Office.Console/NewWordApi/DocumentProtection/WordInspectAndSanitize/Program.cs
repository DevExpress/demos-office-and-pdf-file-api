using DevExpress.Docs.Office;
using DevExpress.Docs.Word;
using System.Text;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const InspectOperationType Operation = InspectOperationType.Inspect;
    const string OutputDirectoryName = "Output";
    const string InspectOutputFileName = "Inspection_Log.txt";
    const string SanitizeOutputFileName = "Sanitized_Document.docx";
    const DocumentFormat SanitizeOutputFormat = DocumentFormat.Docx;
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
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, outputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = Operation == InspectOperationType.Inspect
                ? CreateInspectionLog(inputStream)
                : SanitizeDocument(inputStream, SanitizeOutputFormat, new DocumentSanitizeOptions() {
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

    static Stream CreateInspectionLog(Stream inputStream) {
        var inspectionResult = InspectDocument(inputStream, DocumentInspectOptions.All);
        var outputStream = new MemoryStream();
        using var writer = new StreamWriter(outputStream, Encoding.UTF8, leaveOpen: true);
        writer.Write(inspectionResult);
        writer.Flush();
        outputStream.Position = 0;
        return outputStream;
    }

    static string InspectDocument(Stream inputStream, DocumentInspectOptions inspectOptions) {
        using var document = new WordDocument(inputStream);

        var output = new StringBuilder();

        var inspectResult = document.Inspect(inspectOptions);
        if(inspectResult.ContentTypes.Count == 0) {
            output.AppendLine("No issues found.");
        }
        else {
            output.AppendLine("Document contains the following hidden/sensitive content:");
            output.AppendLine(string.Join(Environment.NewLine, inspectResult.ContentTypes));
        }

        return output.ToString();
    }

    static Stream SanitizeDocument(Stream inputStream, DocumentFormat outputFormat, DocumentSanitizeOptions? sanitizeOptions = null) {
        using var document = new WordDocument(inputStream);

        // Inspect a document and create DocumentSanitizeOptions based on the inspection result
        if(sanitizeOptions == null) {
            var inspectResult = document.Inspect();
            sanitizeOptions = inspectResult.CreateSanitizeOptions();
        }

        // Use the sanitizeResult object to get information about the removed content
        var sanitizeResult = document.Sanitize(sanitizeOptions);

        // Save document
        var outputStream = new MemoryStream();

        document.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }

    enum InspectOperationType {
        Inspect,
        Sanitize
    }
}