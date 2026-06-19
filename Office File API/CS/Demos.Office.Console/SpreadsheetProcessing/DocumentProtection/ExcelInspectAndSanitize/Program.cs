using System.Text;
using DevExpress.Office;
using DevExpress.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputFormat = "txt";
    const InspectOperationType Operation = InspectOperationType.Inspect;
    const string InspectOutputFileName = "Inspection_Log.txt";
    const string SanitizeOutputFileName = "Sanitized_Document.xlsx";
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
            var outputPath = Path.Combine(AppContext.BaseDirectory, outputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = Operation == InspectOperationType.Inspect
                ? CreateInspectionLog(inputStream)
                : SanitizeWorkbook(inputStream, GetDocumentFormat(OutputFormat), new WorkbookSanitizeOptions {
                    HiddenRows = RemoveHiddenRows ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    HiddenColumns = RemoveHiddenColumns ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
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

    static MemoryStream CreateInspectionLog(Stream inputStream) {
        var inspectionResult = InspectWorkbook(inputStream, WorkbookInspectOptions.All);
        var outputStream = new MemoryStream();
        using var writer = new StreamWriter(outputStream, Encoding.UTF8, leaveOpen: true);
        writer.Write(inspectionResult);
        writer.Flush();
        outputStream.Position = 0;
        return outputStream;
    }

    static string InspectWorkbook(Stream inputStream, WorkbookInspectOptions inspectOptions) {
        using var workbook = new Workbook();
        workbook.LoadDocument(inputStream);

        StringBuilder output = new StringBuilder();
        var inspectResult = workbook.Inspect(inspectOptions);
        if(inspectResult.ContentTypes.Count == 0)
            output.AppendLine("No issues found.");
        else {
            output.AppendLine("Document contains the following hidden/sensitive content:");
            output.AppendLine(string.Join(Environment.NewLine, inspectResult.ContentTypes));
        }
        return output.ToString();
    }

    static Stream SanitizeWorkbook(Stream inputStream, DocumentFormat outputFormat, WorkbookSanitizeOptions? sanitizeOptions = null) {
        using var workbook = new Workbook();
        workbook.LoadDocument(inputStream);

        if(sanitizeOptions == null) {
            var inspectResult = workbook.Inspect();
            sanitizeOptions = inspectResult.CreateSanitizeOptions();
        }

        workbook.Sanitize(sanitizeOptions);
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "xls" => DocumentFormat.Xls,
            "xlsx" => DocumentFormat.Xlsx,
            "xlsm" => DocumentFormat.Xlsm,
            "xltx" => DocumentFormat.Xltx,
            "xltm" => DocumentFormat.Xltm,
            "csv" => DocumentFormat.Csv,
            "txt" => DocumentFormat.Text,
            _ => DocumentFormat.Xlsx
        };
    }

    enum InspectOperationType {
        Inspect,
        Sanitize
    }
}
