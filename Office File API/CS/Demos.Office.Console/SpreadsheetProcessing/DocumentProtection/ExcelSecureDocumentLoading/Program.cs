using DevExpress.Office;
using DevExpress.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputFileName = "Secure_Document.xlsx";
    const string OutputFormat = "xlsx";
    const int MaxWorksheetCount = 10;
    const int MaxSheetRowCount = 100000;
    const int MaxSheetCellCount = 1000000;
    const int MaxChartCount = 10;
    const bool RemoveActiveXContent = true;
    const bool RemoveCustomXmlParts = true;
    const bool RemoveExternalConnections = true;
    const bool RemoveExternalWorkbooks = true;
    const bool RemoveMacros = true;
    const bool RemoveOleObjects = true;
    const bool RemovePivotCaches = true;
    const bool RemoveRestrictedFormulas = true;
    const bool RemoveRestrictedLinks = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

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
                    RemoveMacros = RemoveMacros,
                    RemoveOleObjects = RemoveOleObjects,
                    RemovePivotCaches = RemovePivotCaches,
                    RemoveRestrictedFormulas = RemoveRestrictedFormulas,
                    RemoveRestrictedLinks = RemoveRestrictedLinks
                },
                out loadErrorMessage,
                GetDocumentFormat(OutputFormat));
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

    static Stream TryLoadDocument(Stream inputStream, WorkbookSecurityLoadingLimits securityLoadingLimits, WorkbookSecurityLoadingOptions securityLoadingOptions, out string loadErrorMessage, DocumentFormat outputFormat) {
        using var workbook = new Workbook();

        var workbookSecurityLoadingOptions = workbook.Options.SecurityLoadingOptions;
        workbookSecurityLoadingOptions.RemoveActiveXContent = securityLoadingOptions.RemoveActiveXContent;
        workbookSecurityLoadingOptions.RemoveCustomXMLParts = securityLoadingOptions.RemoveCustomXMLParts;
        workbookSecurityLoadingOptions.RemoveExternalConnections = securityLoadingOptions.RemoveExternalConnections;
        workbookSecurityLoadingOptions.RemoveExternalWorkbooks = securityLoadingOptions.RemoveExternalWorkbooks;
        workbookSecurityLoadingOptions.RemoveMacros = securityLoadingOptions.RemoveMacros;
        workbookSecurityLoadingOptions.RemoveOleObjects = securityLoadingOptions.RemoveOleObjects;
        workbookSecurityLoadingOptions.RemovePivotCaches = securityLoadingOptions.RemovePivotCaches;
        workbookSecurityLoadingOptions.RemoveRestrictedFormulas = securityLoadingOptions.RemoveRestrictedFormulas;
        workbookSecurityLoadingOptions.RemoveRestrictedLinks = securityLoadingOptions.RemoveRestrictedLinks;

        var workbookSecurityLoadingLimits = workbook.Options.SecurityLoadingLimits;
        workbookSecurityLoadingLimits.MaxChartCount = securityLoadingLimits.MaxChartCount;
        workbookSecurityLoadingLimits.MaxFileSize = securityLoadingLimits.MaxFileSize;
        workbookSecurityLoadingLimits.MaxSheetCellCount = securityLoadingLimits.MaxSheetCellCount;
        workbookSecurityLoadingLimits.MaxSheetRowCount = securityLoadingLimits.MaxSheetRowCount;
        workbookSecurityLoadingLimits.MaxWorksheetCount = securityLoadingLimits.MaxWorksheetCount;
        workbookSecurityLoadingLimits.MaxXmlElementCount = securityLoadingLimits.MaxXmlElementCount;
        workbookSecurityLoadingLimits.MaxXmlElementDepth = securityLoadingLimits.MaxXmlElementDepth;

        loadErrorMessage = string.Empty;
        try {
            workbook.LoadDocument(inputStream);
            var outputStream = new MemoryStream();
            workbook.SaveDocument(outputStream, outputFormat);
            outputStream.Position = 0;
            return outputStream;
        }
        catch(SecurityLoadingLimitExceededException limitExceededException) {
            loadErrorMessage = limitExceededException.Message;
            return new MemoryStream();
        }
    }

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "xls" => DocumentFormat.Xls,
            "xlsx" => DocumentFormat.Xlsx,
            "xlsm" => DocumentFormat.Xlsm,
            "xltx" => DocumentFormat.Xltx,
            "xltm" => DocumentFormat.Xltm,
            _ => DocumentFormat.Xlsx
        };
    }
}
