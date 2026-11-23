using DevExpress.Office;
using DevExpress.XtraRichEdit;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Secure_Document.docx";
    const string OutputFormat = "docx";
    const int MaxSectionsCount = 10;
    const int MaxParagraphCount = 5000;
    const int MaxTableCount = 50;
    const int MaxTableRowCount = 500;
    const int MaxSubDocumentCount = 200;
    const bool RemoveActiveXContent = true;
    const bool RemoveCustomXmlParts = true;
    const bool RemoveMacros = true;
    const bool RemoveOleObjects = true;
    const RestrictedHyperlinkRemovalMode RestrictedHyperlinkRemovalMode = RestrictedHyperlinkRemovalMode.UriOnly;
    const bool RemoveExternalImages = true;
    const bool RemoveIncludePictureFields = true;
    const bool RemoveDdeFields = true;

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
                new WordProcessingSecurityLoadingLimits {
                    MaxSectionsCount = MaxSectionsCount,
                    MaxParagraphCount = MaxParagraphCount,
                    MaxTableCount = MaxTableCount,
                    MaxTableRowCount = MaxTableRowCount,
                    MaxSubDocumentCount = MaxSubDocumentCount,
                    MaxFileSize = 5 * 1024 * 1024,
                    MaxXmlElementCount = 5000,
                    MaxXmlElementDepth = 50
                },
                new WordProcessingSecurityLoadingOptions {
                    RemoveActiveXContent = RemoveActiveXContent,
                    RemoveCustomXMLParts = RemoveCustomXmlParts,
                    RemoveMacros = RemoveMacros,
                    RemoveOleObjects = RemoveOleObjects,
                    RemoveExternalImages = RemoveExternalImages,
                    RemoveIncludePictureFields = RemoveIncludePictureFields,
                    RemoveDDEFields = RemoveDdeFields,
                    RestrictedHyperlinkRemovalMode = RestrictedHyperlinkRemovalMode
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

    static Stream TryLoadDocument(Stream inputStream, WordProcessingSecurityLoadingLimits securityLoadingLimits, WordProcessingSecurityLoadingOptions securityLoadingOptions, out string loadErrorMessage, DocumentFormat outputFormat) {
        using var richEditDocumentServer = new RichEditDocumentServer();

        var richSecurityLoadingOptions = richEditDocumentServer.Options.SecurityLoadingOptions;
        richSecurityLoadingOptions.RemoveActiveXContent = securityLoadingOptions.RemoveActiveXContent;
        richSecurityLoadingOptions.RemoveCustomXMLParts = securityLoadingOptions.RemoveCustomXMLParts;
        richSecurityLoadingOptions.RemoveMacros = securityLoadingOptions.RemoveMacros;
        richSecurityLoadingOptions.RemoveOleObjects = securityLoadingOptions.RemoveOleObjects;
        richSecurityLoadingOptions.RestrictedHyperlinkRemovalMode = securityLoadingOptions.RestrictedHyperlinkRemovalMode;
        richSecurityLoadingOptions.RemoveExternalImages = securityLoadingOptions.RemoveExternalImages;
        richSecurityLoadingOptions.RemoveIncludePictureFields = securityLoadingOptions.RemoveIncludePictureFields;
        richSecurityLoadingOptions.RemoveDDEFields = securityLoadingOptions.RemoveDDEFields;

        var richSecurityLoadingLimits = richEditDocumentServer.Options.SecurityLoadingLimits;
        richSecurityLoadingLimits.MaxFileSize = securityLoadingLimits.MaxFileSize;
        richSecurityLoadingLimits.MaxXmlElementCount = securityLoadingLimits.MaxXmlElementCount;
        richSecurityLoadingLimits.MaxXmlElementDepth = securityLoadingLimits.MaxXmlElementDepth;
        richSecurityLoadingLimits.MaxSectionsCount = securityLoadingLimits.MaxSectionsCount;
        richSecurityLoadingLimits.MaxParagraphCount = securityLoadingLimits.MaxParagraphCount;
        richSecurityLoadingLimits.MaxTableCount = securityLoadingLimits.MaxTableCount;
        richSecurityLoadingLimits.MaxTableRowCount = securityLoadingLimits.MaxTableRowCount;
        richSecurityLoadingLimits.MaxSubDocumentCount = securityLoadingLimits.MaxSubDocumentCount;

        loadErrorMessage = string.Empty;
        try {
            richEditDocumentServer.LoadDocument(inputStream);
            var outputStream = new MemoryStream();
            richEditDocumentServer.SaveDocument(outputStream, outputFormat);
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
            "doc" => DocumentFormat.Doc,
            "docx" => DocumentFormat.Docx,
            "docm" => DocumentFormat.Docm,
            "dotx" => DocumentFormat.Dotx,
            "dotm" => DocumentFormat.Dotm,
            _ => DocumentFormat.Docx
        };
    }
}
