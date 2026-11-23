using DevExpress.Docs.Ocr;

namespace BlazorDemo.OfficeFileApi.Services {
    #region Code
    public sealed class OcrProviderService : IDisposable {
        readonly Lazy<DocTrOcrProvider> withOrientationDetection = new(() => CreateProvider(true));
        readonly Lazy<DocTrOcrProvider> withoutOrientationDetection = new(() => CreateProvider(false));

        public IOcrProvider GetProvider(bool detectPageOrientation) {
            return detectPageOrientation ? withOrientationDetection.Value : withoutOrientationDetection.Value;
        }

        static DocTrOcrProvider CreateProvider(bool detectPageOrientation) {
            return new DocTrOcrProvider(new DocTrOcrProviderOptions {
                DetectPageOrientation = detectPageOrientation
            });
        }

        public void Dispose() {
            if(withOrientationDetection.IsValueCreated)
                withOrientationDetection.Value.Dispose();
            if(withoutOrientationDetection.IsValueCreated)
                withoutOrientationDetection.Value.Dispose();
        }
    }
    #endregion
}
