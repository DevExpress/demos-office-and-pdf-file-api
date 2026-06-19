using Azure.AI.OpenAI;
using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Docs;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;

internal static class Program {
    const string EndpointUrl = "https://public-api.devexpress.com/demo-openai";
    const string ApiKey = "DEMO";
    const string DeploymentName = "demo";
    static readonly string InputFilePath = Path.Combine("Documents", "LetterTemplate.docx");
    const string OutputFileName = "ResultDocument.docx";
    const SummarizationMode ContentSummarizationMode = SummarizationMode.Extractive;

    static async Task<int> Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

            using var serviceProvider = CreateServiceProvider();
            var service = serviceProvider.GetRequiredService<IAIDocProcessingService>();
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = await SummarizeDocument(service, inputStream, ContentSummarizationMode);
            SaveDocument(outputStream, outputPath);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static ServiceProvider CreateServiceProvider() {
        var services = new ServiceCollection();
        var azureOpenAIClient = new AzureOpenAIClient(new Uri(EndpointUrl), new System.ClientModel.ApiKeyCredential(ApiKey));
        IChatClient chatClient = azureOpenAIClient.GetChatClient(DeploymentName).AsIChatClient();
        services.AddSingleton(chatClient);
        services.AddDevExpressAIConsole(config => config.RegisterAIDocProcessingService());
        return services.BuildServiceProvider();
    }

    static void SaveDocument(Stream stream, string outputPath) {
        stream.Position = 0;
        using var fileStream = File.Create(outputPath);
        stream.CopyTo(fileStream);
    }

    static async Task<Stream> SummarizeDocument(IAIDocProcessingService service, Stream inputStream, SummarizationMode summarizationMode) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(inputStream);

        var result = await service.SummarizeAsync(wordProcessor, summarizationMode);
        var outputStream = new MemoryStream();

        var section = wordProcessor.Document.InsertSection(wordProcessor.Document.Range.Start);
        InsertContentToSection(wordProcessor.Document, result, section);
        wordProcessor.SaveDocument(outputStream, DocumentFormat.OpenXml);

        outputStream.Position = 0;
        return outputStream;
    }

    static void InsertContentToSection(Document document, string text, Section section) {
        CharacterProperties sectionCharacterProperties = document.BeginUpdateCharacters(section.Range);
        sectionCharacterProperties.Reset(CharacterPropertiesMask.All);
        sectionCharacterProperties.Style.Reset(CharacterPropertiesMask.All);
        document.EndUpdateCharacters(sectionCharacterProperties);

        ParagraphProperties sectionParagraphProperties = document.BeginUpdateParagraphs(section.Range);
        sectionParagraphProperties.Reset(ParagraphPropertiesMask.All);
        sectionParagraphProperties.Style.Reset(ParagraphPropertiesMask.All);
        document.EndUpdateParagraphs(sectionParagraphProperties);

        var textRange = document.InsertText(section.Range.Start, "Document Summary");
        document.Paragraphs.Insert(textRange.End);
        document.Paragraphs.Insert(textRange.End);
        document.InsertText(textRange.End, text);

        CharacterProperties titleCharacterProperties = document.BeginUpdateCharacters(section.Paragraphs.First().Range);
        titleCharacterProperties.Bold = true;
        titleCharacterProperties.FontSize = 20;
        document.EndUpdateCharacters(titleCharacterProperties);
    }
}
