using OpenAI;
using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Docs;
using DevExpress.Pdf;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;

internal static class Program {
    const string EndpointUrl = "https://public-api.devexpress.com/demo-openai";
    const string ApiKey = "DEMO";
    const string DeploymentName = "demo";
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "ResultDocument.txt";
    const SummarizationMode ContentSummarizationMode = SummarizationMode.Extractive;

    static async Task<int> Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var serviceProvider = CreateServiceProvider();
            var service = serviceProvider.GetRequiredService<IAIDocProcessingService>();
            using var inputStream = File.OpenRead(inputPath);
            var result = await SummarizeDocument(service, inputStream, ContentSummarizationMode);
            File.WriteAllText(outputPath, result);

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
        var azureOpenAIV1Endpoint = new Uri($"{EndpointUrl.TrimEnd('/')}/openai/v1");
        var azureOpenAIClient = new OpenAIClient(
            new System.ClientModel.ApiKeyCredential(ApiKey),
            new OpenAIClientOptions() { Endpoint = azureOpenAIV1Endpoint }
        );
        IChatClient chatClient = azureOpenAIClient.GetChatClient(DeploymentName).AsIChatClient();
        services.AddSingleton(chatClient);
        services.AddDevExpressAIConsole(config => config.RegisterAIDocProcessingService());
        return services.BuildServiceProvider();
    }

    static async Task<string> SummarizeDocument(IAIDocProcessingService service, Stream inputStream, SummarizationMode summarizationMode) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(inputStream);
        return await service.SummarizeAsync(processor, summarizationMode);
    }
}
