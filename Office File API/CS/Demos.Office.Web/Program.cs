using BlazorDemo.OfficeFileApi;
using BlazorDemo.OfficeFileApi.DemoShell;
using BlazorDemo.OfficeFileApi.Services;
using BlazorDemo.Services;
using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Docs;
using DevExpress.Blazor.DocumentMetadata;
using Microsoft.Extensions.AI;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();
builder.Services.AddDevExpressBlazor(options => {
    options.SizeMode = DevExpress.Blazor.SizeMode.Medium;
});
builder.Services.AddSingleton<InMemoryNWindDataService>();
builder.Services.AddSingleton<FileFormatService>();
builder.Services.AddSingleton<SpreadsheetProtectionOptionsService>();
builder.Services.AddSingleton<OcrProviderService>();
builder.Services.AddSingleton<IOfficeFileApiBlazorDemoSectionSvgService, SvgService>();
builder.Services.AddOptions();
builder.Services.AddHttpContextAccessor();
builder.Services.AddSingleton<DemoConfiguration>();
builder.Services.AddScoped<DemoThemeService>();
builder.Services.AddScoped<IDemoStaticResourceService, DemoStaticResourceService>();
builder.Services.AddDocumentMetadata((sp, metadataCollection) => { sp!.GetService<DemoConfiguration>()!.ConfigureMetadata(metadataCollection); });
builder.Services.AddSingleton<IDemoVersion, DemoVersion>(x => {
    var dxVersion = new Version(AssemblyInfo.Version);
    return new DemoVersion(new Version(dxVersion.Major, dxVersion.Minor, dxVersion.Build).ToString());
});

var azureOpenAIEndpoint = builder.Configuration.GetSection("AIIntegrationSettings")["EndpointUrl"];
var azureOpenAIKey = builder.Configuration.GetSection("AIIntegrationSettings")["Key"];
var deploymentName = builder.Configuration.GetSection("AIIntegrationSettings")["DeploymentName"];

// Create an Azure OpenAI client with endpoint and API key from helper.
var azureOpenAIV1Endpoint = new Uri($"{azureOpenAIEndpoint.TrimEnd('/')}/openai/v1");
var azureOpenAIClient = new OpenAI.OpenAIClient(
    new System.ClientModel.ApiKeyCredential(azureOpenAIKey),
    new OpenAI.OpenAIClientOptions() { Endpoint = azureOpenAIV1Endpoint }
);

// Get a model-specific chat client and adapt it to IChatClient.
IChatClient chatClient = azureOpenAIClient
    .GetChatClient(deploymentName)
    .AsIChatClient();

// Register the chat client as a singleton in the dependency injection container.
builder.Services.AddSingleton(chatClient);

// Add DevExpress AI services and register the document-processing extensions.
builder.Services.AddDevExpressAIConsole((config) => {
    config.RegisterAIDocProcessingService();
});

// Add HTTP client for API calls
builder.Services.AddHttpClient();

// Add MVC services with controllers - REQUIRED for API controllers to work
builder.Services.AddControllers();

builder.WebHost.UseWebRoot("wwwroot");
builder.WebHost.UseStaticWebAssets();

var app = builder.Build();
if(!app.Environment.IsDevelopment()) {
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseStatusCodePagesWithReExecute("/");
app.UseAntiforgery();
app.MapControllers(); // Map controller routes
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();
app.Run();
