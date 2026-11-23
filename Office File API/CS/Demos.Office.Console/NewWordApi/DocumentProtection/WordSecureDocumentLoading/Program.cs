using DevExpress.Docs;
using DevExpress.Docs.Office;
using DevExpress.Docs.Word;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Secure_Document.docx";
    const DocumentFormat OutputFormat = DocumentFormat.Docx;
    const int MaxSectionsCount = 10;
    const int MaxParagraphCount = 500;
    const int MaxTableCount = 10;
    const int MaxTableRowCount = 100;
    const int MaxSubDocumentCount = 50;
    const bool RemoveActiveXContent = true;
    const bool RemoveCustomXmlParts = true;
    const bool RemoveMacros = true;
    const bool RemoveOleObjects = true;
    const bool RemoveExternalImages = true;
    const bool RemoveIncludePictureFields = true;
    const bool RemoveDdeFields = true;
    const RestrictedHyperlinkRemovalMode HyperlinkRemovalMode = RestrictedHyperlinkRemovalMode.UriOnly;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = TryLoadDocument(
                inputStream,
                new DocumentSecurityLoadingLimits() {
                    MaxSectionsCount = MaxSectionsCount,
                    MaxParagraphCount = MaxParagraphCount,
                    MaxTableCount = MaxTableCount,
                    MaxTableRowCount = MaxTableRowCount,
                    MaxSubDocumentCount = MaxSubDocumentCount,
                    MaxFileSize = 5 * 1024 * 1024,
                    MaxXmlElementCount = 5000,
                    MaxXmlElementDepth = 50
                },
                new DocumentSecurityLoadingOptions() {
                    RemoveActiveXContent = RemoveActiveXContent,
                    RemoveCustomXMLParts = RemoveCustomXmlParts,
                    RemoveMacros = RemoveMacros,
                    RemoveOleObjects = RemoveOleObjects,
                    RemoveExternalImages = RemoveExternalImages,
                    RemoveIncludePictureFields = RemoveIncludePictureFields,
                    RemoveDDEFields = RemoveDdeFields,
                    RestrictedHyperlinkRemovalMode = HyperlinkRemovalMode
                },
                out var loadErrorMessage,
                OutputFormat);

            if(!string.IsNullOrEmpty(loadErrorMessage)) {
                throw new InvalidOperationException(loadErrorMessage);
            }

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

    static Stream TryLoadDocument(Stream inputStream, DocumentSecurityLoadingLimits securityLoadingLimits,
                                  DocumentSecurityLoadingOptions securityLoadingOptions, out string loadErrorMessage, DocumentFormat outputFormat) {
        var loadOptions = new LoadOptions() {
            SecurityLoadingHandler = new DocumentSecurityLoadingHandler(),
            SecurityLoadingOptions = securityLoadingOptions,
            SecurityLoadingLimits = securityLoadingLimits
        };

        loadErrorMessage = "";
        try {
            using var document = new WordDocument(inputStream, loadOptions);

            // Save document
            var outputStream = new MemoryStream();

            document.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        catch(SecurityLoadingLimitExceededException limitExceededException) {
            loadErrorMessage = limitExceededException.Message;
            return new MemoryStream();
        }
    }

    class DocumentSecurityLoadingHandler : ISecurityLoadingHandler {
        public SecurityLimitAction HandleLimitExceeded(SecurityLimitExceededParams parameters) {
            // Use the HandleLimitExceeded method to log information about the security loading limits
            // Return SecurityLimitAction.Continue to load a document even if it exceeds the specified limits
            return SecurityLimitAction.Throw;
        }

        public SecurityViolationAction HandleOptionsViolation(SecurityOptionsViolationParams parameters) {
            // Use the HandleOptionsViolation method to log information about the security loading options violation
            // Return SecurityViolationAction.Keep to ignore the violation and load a document even if it contains unsafe content
            return SecurityViolationAction.Remove;
        }
    }
}