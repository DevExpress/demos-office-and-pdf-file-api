using System.Drawing;
using System.Security.Cryptography.X509Certificates;
using DevExpress.Drawing;
using DevExpress.Docs.Pdf;
using DevExpress.Office.Tsp;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Unsigned.pdf");
    static readonly string CertificatePath = Path.Combine("Documents", "PDF", "Sample_Certificate.pfx");
    static readonly string SignatureImagePath = Path.Combine("Documents", "DevExpress.png");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Document.pdf";
    const string CertificatePassword = "dxdemo";
    const string TsaUrl = "https://www.freetsa.org/tsr";
    const HashAlgorithm CertificateHashAlgorithm = HashAlgorithm.SHA256;
    const DevExpress.Office.DigitalSignatures.HashAlgorithmType TsaHashAlgorithm = DevExpress.Office.DigitalSignatures.HashAlgorithmType.SHA256;
    const bool ShowDate = true;
    const bool ShowLocation = true;
    const bool ShowName = true;
    const bool ShowReason = true;
    const bool UseCustomAppearance = false;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var certificatePath = Path.Combine(AppContext.BaseDirectory, CertificatePath);
            var signatureImagePath = Path.Combine(AppContext.BaseDirectory, SignatureImagePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var documentStream = File.OpenRead(inputPath);
            using var certificateStream = File.OpenRead(certificatePath);
            using var signatureImageStream = File.OpenRead(signatureImagePath);

            using var outputStream = Sign(
                documentStream,
                certificateStream,
                signatureImageStream,
                CertificateHashAlgorithm,
                TsaHashAlgorithm,
                UseCustomAppearance
                    ? new AppearanceSettings {
                        ShowDate = ShowDate,
                        ShowLocation = ShowLocation,
                        ShowName = ShowName,
                        ShowReason = ShowReason
                    }
                    : null);

            SaveDocument(outputStream, outputPath);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static void SaveDocument(Stream stream, string outputPath) {
        stream.Position = 0;
        using var fileStream = File.Create(outputPath);
        stream.CopyTo(fileStream);
    }

    static MemoryStream Sign(
        Stream documentStream,
        Stream certificateStream,
        Stream signatureImageStream,
        HashAlgorithm certificateHashAlgorithm,
        DevExpress.Office.DigitalSignatures.HashAlgorithmType tsaHashAlgorithm,
        AppearanceSettings? appearanceSettings = null) {
        using var doc = new PdfDocument(documentStream);

        var signatureField = CreateSignatureField(doc);
        var signatureWidget = CreateSignatureWidget(doc.Pages.Last(), signatureField);
        using var certificate = LoadCertificate(certificateStream, CertificatePassword);

        var signatureOptions = new SignatureOptions {
            Certificate = certificate,
            CertificationLevel = CertificationLevel.NoCertification,
            ContactInfo = "support@devexpress.com",
            HashAlgorithm = certificateHashAlgorithm,
            Location = "US",
            SignerName = "John Doe",
            Reason = "Demo",
            SigningTime = DateTime.UtcNow
        };

        using var dxImage = DXImage.FromStream(signatureImageStream);

        if(appearanceSettings == null)
            BuildSignatureAppearance(signatureWidget, dxImage);
        else
            BuildSignatureAppearance(signatureWidget, dxImage, signatureOptions, appearanceSettings);

        var tsaClient = new TsaClient(new Uri(TsaUrl, UriKind.Absolute), tsaHashAlgorithm);

        var documentSigner = new Pkcs7DocumentSigner(signatureOptions, tsaClient);
        doc.AddSignature(signatureField, documentSigner);

        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }

    static SignatureField CreateSignatureField(PdfDocument document) {
        var signatureField = new SignatureField("SignatureField");
        document.Fields.Add(signatureField);
        return signatureField;
    }

    static SignatureWidgetAnnotation CreateSignatureWidget(Page page, SignatureField signatureField) {
        var signatureFieldBounds = new RectangleF(226, 50, 160, 30);
        var signatureFieldAnnotation = new SignatureWidgetAnnotation(signatureField, signatureFieldBounds);
        page.Annotations.Add(signatureFieldAnnotation);
        return signatureFieldAnnotation;
    }

    static X509Certificate2 LoadCertificate(Stream certificateStream, string password) {
        using MemoryStream buffer = new();
        certificateStream.CopyTo(buffer);
        return X509CertificateLoader.LoadPkcs12(buffer.ToArray(), password);
    }

    static void BuildSignatureAppearance(SignatureWidgetAnnotation widget, DXImage image) {
        float boxWidth = widget.Bounds.Width;
        float boxHeight = widget.Bounds.Height;
        float scale = Math.Min((boxWidth - 4) / image.Width, (boxHeight - 4) / image.Height);
        float displayWidth = image.Width * scale;
        float displayHeight = image.Height * scale;

        ImageFragment imageFragment = widget.AddImageFragment(image, 2, (boxHeight - displayHeight) / 2);
        imageFragment.ScaleX = displayWidth;
        imageFragment.ScaleY = displayHeight;
    }

    static void BuildSignatureAppearance(SignatureWidgetAnnotation widget, DXImage image, SignatureOptions signatureOptions, AppearanceSettings appearanceSettings) {
        float boxWidth = widget.Bounds.Width;
        float boxHeight = widget.Bounds.Height;
        float imageAreaWidth = boxWidth / 3;
        float scale = Math.Min((imageAreaWidth - 4) / image.Width, (boxHeight - 4) / image.Height);
        float displayWidth = image.Width * scale;
        float displayHeight = image.Height * scale;

        ImageFragment imageFragment = widget.AddImageFragment(image, 2, (boxHeight - displayHeight) / 2);
        imageFragment.ScaleX = displayWidth;
        imageFragment.ScaleY = displayHeight;

        float fontSize = 6;
        float lineHeight = 7;
        float textX = imageAreaWidth + 2;
        float textY = boxHeight - lineHeight;
        if(appearanceSettings.ShowName) {
            widget.AddTextFragment($"Digitally signed by {signatureOptions.SignerName}", textX, textY).FontSize = fontSize;
            textY -= lineHeight;
        }
        if(appearanceSettings.ShowReason) {
            widget.AddTextFragment($"Reason: {signatureOptions.Reason}", textX, textY).FontSize = fontSize;
            textY -= lineHeight;
        }
        if(appearanceSettings.ShowLocation) {
            widget.AddTextFragment($"Location: {signatureOptions.Location}", textX, textY).FontSize = fontSize;
            textY -= lineHeight;
        }
        if(appearanceSettings.ShowDate)
            widget.AddTextFragment($"Date: {signatureOptions.SigningTime?.ToString("yyyy'.'MM'.'dd HH':'mm':'ss zzz")}", textX, textY).FontSize = fontSize;
    }

    class AppearanceSettings {
        public bool ShowDate { get; set; }
        public bool ShowLocation { get; set; }
        public bool ShowName { get; set; }
        public bool ShowReason { get; set; }
    }
}
