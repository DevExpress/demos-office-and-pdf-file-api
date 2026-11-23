using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Rendering;
using Microsoft.AspNetCore.Components.Routing;

namespace DevExpress.Blazor.DocumentMetadata {
    public sealed class DocumentMetadataHeadContent : ComponentBase, IDisposable {
        [Inject]
        public IDocumentMetadataRendererSource RendererSource { get; set; }
        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        protected override void OnInitialized() {
            NavigationManager.LocationChanged += OnLocationChanged;
        }

        void OnLocationChanged(object sender, LocationChangedEventArgs args) {
            _ = InvokeAsync(StateHasChanged);
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "ASP0006:Do not use non-literal sequence numbers", Justification = "<Pending>")]
        protected override void BuildRenderTree(RenderTreeBuilder builder) {
            var renderers = RendererSource.GetRenderers();
            foreach (var renderer in renderers) {
                renderer.Render(builder, 0, NavigationManager);
            }
        }

        void IDisposable.Dispose() {
            NavigationManager.LocationChanged -= OnLocationChanged;
        }
    }
}
