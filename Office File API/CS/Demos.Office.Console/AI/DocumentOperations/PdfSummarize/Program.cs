using Azure.AI.OpenAI;
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
    const string OutputFileName = "ResultDocument.txt";
    const SummarizationMode ContentSummarizationMode = SummarizationMode.Extractive;

    static async Task<int> Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

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
        var azureOpenAIClient = new AzureOpenAIClient(new Uri(EndpointUrl), new System.ClientModel.ApiKeyCredential(ApiKey));
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
