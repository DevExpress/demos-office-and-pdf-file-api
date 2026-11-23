using DevExpress.Docs;
using DevExpress.Docs.Office;
using DevExpress.Docs.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Secure_Document.xlsx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;
    const int MaxWorksheetCount = 1;
    const int MaxSheetRowCount = 100;
    const int MaxSheetCellCount = 500;
    const int MaxChartCount = 1;
    const bool RemoveActiveXContent = true;
    const bool RemoveCustomXmlParts = true;
    const bool RemoveExternalConnections = true;
    const bool RemoveExternalWorkbooks = true;
    const bool RemoveExternalImages = true;
    const bool RemoveMacros = true;
    const bool RemoveOleObjects = true;
    const bool RemovePivotCaches = true;
    const bool RemoveRestrictedFormulas = true;
    const RestrictedHyperlinkRemovalMode RestrictedHyperlinkRemovalMode = DevExpress.Docs.Office.RestrictedHyperlinkRemovalMode.UriOnly;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            string loadErrorMessage;
            using var outputStream = TryLoadDocument(
                inputStream,
                new WorkbookSecurityLoadingLimits {
                    MaxChartCount = MaxChartCount,
                    MaxFileSize = 5 * 1024 * 1024,
                    MaxSheetCellCount = MaxSheetCellCount,
                    MaxSheetRowCount = MaxSheetRowCount,
                    MaxWorksheetCount = MaxWorksheetCount,
                    MaxXmlElementCount = 1000,
                    MaxXmlElementDepth = 10
                },
                new WorkbookSecurityLoadingOptions {
                    RemoveActiveXContent = RemoveActiveXContent,
                    RemoveCustomXMLParts = RemoveCustomXmlParts,
                    RemoveExternalConnections = RemoveExternalConnections,
                    RemoveExternalWorkbooks = RemoveExternalWorkbooks,
                    RemoveExternalImages = RemoveExternalImages,
                    RemoveMacros = RemoveMacros,
                    RemoveOleObjects = RemoveOleObjects,
                    RemovePivotCaches = RemovePivotCaches,
                    RemoveRestrictedFormulas = RemoveRestrictedFormulas,
                    RestrictedHyperlinkRemovalMode = RestrictedHyperlinkRemovalMode
                },
                out loadErrorMessage,
                OutputFormat);
            if(!string.IsNullOrEmpty(loadErrorMessage))
                throw new InvalidOperationException(loadErrorMessage);

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

    static Stream TryLoadDocument(Stream inputStream, WorkbookSecurityLoadingLimits securityLoadingLimits,
        WorkbookSecurityLoadingOptions securityLoadingOptions, out string loadErrorMessage, DocumentFormat outputFormat) {
        var loadOptions = new LoadOptions() {
            SecurityLoadingHandler = new WorkbookSecurityLoadingHandler(),
            SecurityLoadingOptions = securityLoadingOptions,
            SecurityLoadingLimits = securityLoadingLimits
        };

        loadErrorMessage = "";
        try {
            using var workbook = new Workbook(inputStream, loadOptions);

            var outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        catch(SecurityLoadingLimitExceededException limitExceededException) {
            loadErrorMessage = limitExceededException.Message;
            return new MemoryStream();
        }
    }

    class WorkbookSecurityLoadingHandler : ISecurityLoadingHandler {
        public SecurityLimitAction HandleLimitExceeded(SecurityLimitExceededParams parameters) {
            return SecurityLimitAction.Throw;
        }

        public SecurityViolationAction HandleOptionsViolation(SecurityOptionsViolationParams parameters) {
            return SecurityViolationAction.Remove;
        }
    }
}
