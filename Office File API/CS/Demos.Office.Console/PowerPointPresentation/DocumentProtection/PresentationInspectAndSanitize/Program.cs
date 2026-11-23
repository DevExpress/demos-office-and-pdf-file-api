using DevExpress.Docs.Presentation;
using DevExpress.Docs.Office;
using System.Text;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.pptx");
    const string OutputFormat = "pptx";
    const InspectOperationType Operation = InspectOperationType.Inspect;
    const string OutputDirectoryName = "Output";
    const string InspectOutputFileName = "Inspection_Log.txt";
    const string SanitizeOutputFileName = "Sanitized_Presentation.pptx";
    const bool RemoveHiddenSlides = true;
    const bool RemoveHiddenShapes = true;
    const bool RemoveActiveXContent = true;
    const bool RemoveComments = true;
    const bool RemoveCustomXmlParts = true;
    const bool RemoveEmbeddedChartData = true;
    const bool RemoveMacros = true;
    const bool RemoveMetadata = true;
    const bool RemoveNotes = true;
    const bool RemoveOleObjects = true;
    const bool RemoveOffSlideContent = true;

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
                : SanitizePresentation(inputStream, GetDocumentFormat(OutputFormat), new PresentationSanitizeOptions() {
                    HiddenSlides = RemoveHiddenSlides ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    HiddenShapes = RemoveHiddenShapes ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    RemoveActiveXContent = RemoveActiveXContent,
                    RemoveComments = RemoveComments,
                    RemoveCustomXmlParts = RemoveCustomXmlParts,
                    RemoveEmbeddedChartData = RemoveEmbeddedChartData,
                    RemoveMacros = RemoveMacros,
                    Metadata = RemoveMetadata ? MetadataRemovalScope.All : MetadataRemovalScope.None,
                    RemoveNotes = RemoveNotes,
                    RemoveOleObjects = RemoveOleObjects,
                    RemoveOffSlideContent = RemoveOffSlideContent
                });
            using var fileStream = File.Create(outputPath);
            outputStream.CopyTo(fileStream);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static MemoryStream CreateInspectionLog(Stream inputStream) {
        var inspectionResult = InspectPresentation(inputStream, PresentationInspectOptions.All);
        var outputStream = new MemoryStream();
        using var writer = new StreamWriter(outputStream, Encoding.UTF8, leaveOpen: true);
        writer.Write(inspectionResult);
        writer.Flush();
        outputStream.Position = 0;
        return outputStream;
    }

    static string InspectPresentation(Stream inputStream, PresentationInspectOptions inspectOptions) {
        using var presentation = new Presentation(inputStream);

        StringBuilder output = new StringBuilder();

        var inspectResult = presentation.Inspect(inspectOptions);
        if (inspectResult.ContentTypes.Count == 0) {
            output.AppendLine("No issues found.");
        } else {
            output.AppendLine("Presentation contains the following hidden/sensitive content:");
            output.AppendLine(string.Join(Environment.NewLine, inspectResult.ContentTypes));
        }
        return output.ToString();
    }

    static Stream SanitizePresentation(Stream inputStream, DocumentFormat outputFormat, PresentationSanitizeOptions sanitizeOptions = null) {
        using var presentation = new Presentation(inputStream);

        if (sanitizeOptions == null) {
            var inspectResult = presentation.Inspect();
            sanitizeOptions = inspectResult.CreateSanitizeOptions();
        }

        presentation.Sanitize(sanitizeOptions);
        var outputStream = new MemoryStream();
        presentation.Save(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "potm" => DocumentFormat.Potm,
            "potx" => DocumentFormat.Potx,
            "pptm" => DocumentFormat.Pptm,
            "pptx" => DocumentFormat.Pptx,
            _ => DocumentFormat.Pptx
        };
    }

    enum InspectOperationType {
        Inspect,
        Sanitize
    }
}
