using System.Text;
using DevExpress.Docs.Office;
using DevExpress.Docs.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const InspectOperationType Operation = InspectOperationType.Inspect;
    const string OutputDirectoryName = "Output";
    const string InspectOutputFileName = "Inspection_Log.txt";
    const string SanitizeOutputFileName = "Sanitized_Document.xlsx";
    static readonly DocumentFormat SanitizeOutputFormat = DocumentFormat.Xlsx;
    const bool RemoveHiddenRows = true;
    const bool RemoveHiddenColumns = true;
    const bool RemoveHiddenSheets = true;
    const bool RemoveInvisibleCellText = true;
    const bool RemoveMetadata = true;
    const bool RemoveComments = true;
    const bool RemoveCustomXmlParts = true;
    const bool RemoveActiveXContent = true;
    const bool RemoveMacros = true;
    const bool RemoveOleObjects = true;
    const bool RemoveThreadedComments = true;

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
                : SanitizeWorkbook(inputStream, SanitizeOutputFormat, new WorkbookSanitizeOptions {
                    HiddenColumns = RemoveHiddenColumns ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    HiddenRows = RemoveHiddenRows ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    HiddenSheets = RemoveHiddenSheets ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    InvisibleCellText = RemoveInvisibleCellText ? InvisibleContentSanitizeMode.Remove : InvisibleContentSanitizeMode.Ignore,
                    RemoveMacros = RemoveMacros,
                    RemoveActiveXContent = RemoveActiveXContent,
                    RemoveOleObjects = RemoveOleObjects,
                    RemoveThreadedComments = RemoveThreadedComments,
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
        string inspectionResult = InspectWorkbook(inputStream, WorkbookInspectOptions.All);
        var outputStream = new MemoryStream();
        using var writer = new StreamWriter(outputStream, Encoding.UTF8, leaveOpen: true);
        writer.Write(inspectionResult);
        writer.Flush();
        outputStream.Position = 0;
        return outputStream;
    }

    static string InspectWorkbook(Stream inputStream, WorkbookInspectOptions inspectOptions) {
        using var workbook = new Workbook(inputStream);

        StringBuilder output = new StringBuilder();

        var inspectResult = workbook.Inspect(inspectOptions);
        if(inspectResult.ContentTypes.Count == 0) {
            output.AppendLine("No issues found.");
        }
        else {
            output.AppendLine("Document contains the following hidden/sensitive content:");
            output.AppendLine(string.Join(Environment.NewLine, inspectResult.ContentTypes));
        }
        return output.ToString();
    }

    static Stream SanitizeWorkbook(Stream inputStream, DocumentFormat outputFormat, WorkbookSanitizeOptions? sanitizeOptions = null) {
        using var workbook = new Workbook(inputStream);

        if(sanitizeOptions == null) {
            var inspectResult = workbook.Inspect();
            sanitizeOptions = inspectResult.CreateSanitizeOptions();
        }

        var sanitizeResult = workbook.Sanitize(sanitizeOptions);

        var outputStream = new MemoryStream();

        workbook.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }

    enum InspectOperationType {
        Inspect,
        Sanitize
    }
}
