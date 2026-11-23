using DevExpress.Docs;
using DevExpress.Docs.Presentation;
using RestrictedHyperlinkRemovalMode = DevExpress.Docs.Office.RestrictedHyperlinkRemovalMode;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.pptx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Secure_Document.pptx";
    const string OutputFormat = "pptx";
    const int MaxSlideCount = 10;
    const int MaxShapeCountPerSlide = 20;
    const bool RemoveActiveXContent = true;
    const bool RemoveCustomXmlParts = true;
    const bool RemoveEmbeddedChartData = true;
    const bool RemoveExternalImages = true;
    const bool RemoveMacros = true;
    const bool RemoveOleObjects = true;
    const RestrictedHyperlinkRemovalMode RestrictedHyperlinkRemovalMode = RestrictedHyperlinkRemovalMode.UriOnly;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            string loadErrorMessage;
            using var outputStream = TryLoadPresentation(
                inputStream,
                new PresentationSecurityLoadingLimits() {
                    MaxFileSize = 5 * 1024 * 1024,
                    MaxShapeCountPerSlide = MaxShapeCountPerSlide,
                    MaxSlideCount = MaxSlideCount,
                    MaxXmlElementCount = 1000,
                    MaxXmlElementDepth = 20
                },
                new PresentationSecurityLoadingOptions() {
                    RemoveActiveXContent = RemoveActiveXContent,
                    RemoveCustomXMLParts = RemoveCustomXmlParts,
                    RemoveEmbeddedChartData = RemoveEmbeddedChartData,
                    RemoveExternalImages = RemoveExternalImages,
                    RemoveMacros = RemoveMacros,
                    RemoveOleObjects = RemoveOleObjects,
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
    static Stream TryLoadPresentation(Stream inputStream, PresentationSecurityLoadingLimits securityLoadingLimits, PresentationSecurityLoadingOptions securityLoadingOptions, out string loadErrorMessage, DocumentFormat outputFormat) {
        var loadOptions = new LoadOptions() {
            SecurityLoadingOptions = securityLoadingOptions,
            SecurityLoadingLimits = securityLoadingLimits
        };

        loadErrorMessage = string.Empty;
        try {
            using var presentation = new Presentation(inputStream, loadOptions);
            var outputStream = new MemoryStream();
            presentation.Save(outputStream, outputFormat);
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
            "potm" => DocumentFormat.Potm,
            "potx" => DocumentFormat.Potx,
            "pptm" => DocumentFormat.Pptm,
            "pptx" => DocumentFormat.Pptx,
            _ => DocumentFormat.Pptx
        };
    }
}
