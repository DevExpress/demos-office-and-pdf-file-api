using System.Globalization;
using Azure.AI.OpenAI;
using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Docs;
using DevExpress.Docs.Presentation;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;

internal static class Program {
    const string EndpointUrl = "https://public-api.devexpress.com/demo-openai";
    const string ApiKey = "DEMO";
    const string DeploymentName = "demo";
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.pptx");
    const string OutputFileName = "ResultPresentation.pptx";
    const string TargetLanguage = "de-DE";

    static async Task<int> Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

            using var serviceProvider = CreateServiceProvider();
            var service = serviceProvider.GetRequiredService<IAIDocProcessingService>();
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = await TranslateDocument(service, inputStream, TargetLanguage);
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

    static async Task<Stream> TranslateDocument(IAIDocProcessingService service, Stream inputStream, string targetLanguage) {
        using var presentation = new Presentation(inputStream);
        await service.TranslateAsync(presentation, new CultureInfo(targetLanguage));
        var outputStream = new MemoryStream();
        presentation.SaveDocument(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
