using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Rendering;

namespace DevExpress.Blazor.DocumentMetadata {
    public sealed class DocumentMetadataHeadContent : ComponentBase, IObserver<Renderer>, IDisposable {
        readonly List<Renderer> _renderers = new();
        IDisposable _subscription;
        bool _initialized;

        [Inject]
        public IObservable<Renderer> Metadata { get; set; }
        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        protected override void OnInitialized() {
            _subscription = Metadata!.Subscribe(this);
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "ASP0006:Do not use non-literal sequence numbers", Justification = "<Pending>")]
        protected override void BuildRenderTree(RenderTreeBuilder builder) {
            _renderers.ForEach(r => r.Render(builder, 0, NavigationManager));
        }

        void IObserver<Renderer>.OnNext(Renderer value) {
            if(_initialized) {
                _renderers.Clear();
                _initialized = false;
                _ = InvokeAsync(() => {
                    _initialized = true;
                    StateHasChanged();
                });
            }
            _renderers.Add(value);
        }

        void IObserver<Renderer>.OnCompleted() {
            _initialized = true;
            _ = InvokeAsync(StateHasChanged);
        }

        void IObserver<Renderer>.OnError(Exception error) { }

        void IDisposable.Dispose() => _subscription?.Dispose();
    }
}
