using Azure.AI.OpenAI;
using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Docs;
using DevExpress.Docs.Office;
using DevExpress.Docs.Presentation;
using DevExpress.Office.Utils;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;

internal static class Program {
    const string EndpointUrl = "https://public-api.devexpress.com/demo-openai";
    const string ApiKey = "DEMO";
    const string DeploymentName = "demo";
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.pptx");
    const string OutputFileName = "ResultPresentation.pptx";
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
        using var presentation = new Presentation(inputStream);
        var result = await service.SummarizeAsync(presentation, summarizationMode);
        var outputStream = new MemoryStream();
        var slide = CreateSummarySlide(presentation, result);
        presentation.Slides.Insert(0, slide);
        presentation.SaveDocument(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }

    static Slide CreateSummarySlide(Presentation presentation, string text) {
        SlideMaster slideMaster = presentation.SlideMasters[0];
        SlideLayout slideLayout = slideMaster.Layouts.GetOrCreate(SlideLayoutType.Object);
        Slide slide = new(slideLayout);
        foreach(Shape shape in slide.Shapes) {
            if(shape.PlaceholderSettings.Type is PlaceholderType.Title)
                shape.TextArea = new TextArea("Document Summary");
            if(shape.PlaceholderSettings.Type is PlaceholderType.Body) {
                shape.TextArea = new TextArea(text);
                foreach(var paragraph in shape.TextArea.Paragraphs)
                    paragraph.Properties.ListBullet = ListBullet.None;
            }
        }
        return slide;
    }
}
