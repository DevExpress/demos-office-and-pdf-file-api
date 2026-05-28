namespace DevExpress.Blazor.DocumentMetadata {
    public static partial class DocumentMetadataExtensions {

        public static void AddDocumentMetadata(this IServiceCollection services, Action<IServiceProvider, IDocumentMetadataCollection> initialize) {
            services.AddSingleton(new DocumentMetadataSetup(initialize));
            services.AddDocumentMetadata();
        }
    }
}
