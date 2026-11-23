using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentProtection;
using Microsoft.AspNetCore.Mvc;

#region Code
using DevExpress.Office.Tsp;
using DevExpress.Docs.Pdf;
using System.Drawing;
using System.Security.Cryptography.X509Certificates;
using DevExpress.Drawing;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.DocumentProtection {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfDigitalSignatureController : PdfControllerBase {
        [HttpPost("document-protection/sign")]
        public async Task<IActionResult> SignPdfWithCertificate(
            IFormFile file,
            IFormFile certificate,
            IFormFile signatureImage,
            [FromForm] PdfDigitalSignatureModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var certificateStream = new MemoryStream();
            await certificate.CopyToAsync(certificateStream);
            certificateStream.Position = 0;

            var signatureAppearanceImageStream = new MemoryStream();
            await signatureImage.CopyToAsync(signatureAppearanceImageStream);
            signatureAppearanceImageStream.Position = 0;

            var outputStream = Sign(
                documentStream,
                certificateStream,
                signatureAppearanceImageStream,
                model.CertificateHashAlgorithm,
                model.TsaHashAlgorithm,
                model.UseCustomAppearance ? new AppearanceSettings { ShowDate = model.ShowDate, ShowLocation = model.ShowLocation, ShowName = model.ShowName, ShowReason = model.ShowReason } : null);

            outputStream.Position = 0;

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        MemoryStream Sign(Stream documentStream, Stream certificateStream, Stream signatureImageStream,
            HashAlgorithm certificateHashAlgorithm, DevExpress.Office.DigitalSignatures.HashAlgorithmType tsaHashAlgorithm,
            AppearanceSettings appearanceSettings = null) {
            using var doc = new PdfDocument(documentStream);

            var signatureField = CreateSignatureField(doc);
            var signatureWidget = CreateSignatureWidget(doc.Pages.Last(), signatureField);
            using var certificate = LoadCertificate(certificateStream, "dxdemo");

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

            var tsaClient = new TsaClient(
                new Uri("https://www.freetsa.org/tsr", UriKind.Absolute),
                tsaHashAlgorithm);

            var documentSigner = new Pkcs7DocumentSigner(signatureOptions, tsaClient);            
            doc.AddSignature(signatureField, documentSigner);

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        SignatureField CreateSignatureField(PdfDocument document) {
            var signatureField = new SignatureField("SignatureField");
            document.Fields.Add(signatureField);
            return signatureField;
        }
        SignatureWidgetAnnotation CreateSignatureWidget(Page page, SignatureField signatureField) {
            var signatureFieldBounds = new RectangleF(226, 50, 160, 30);
            var signatureFieldAnnotation = new SignatureWidgetAnnotation(signatureField, signatureFieldBounds);
            page.Annotations.Add(signatureFieldAnnotation);
            return signatureFieldAnnotation;
        }
        X509Certificate2 LoadCertificate(Stream certificateStream, string password) {
            using MemoryStream buffer = new MemoryStream();
            certificateStream.CopyTo(buffer);
            return X509CertificateLoader.LoadPkcs12(buffer.ToArray(), password);
        }
        void BuildSignatureAppearance(SignatureWidgetAnnotation widget, DXImage image) {
            float boxWidth = widget.Bounds.Width;
            float boxHeight = widget.Bounds.Height;
            float scale = Math.Min((boxWidth - 4) / image.Width, (boxHeight - 4) / image.Height);
            float displayWidth = image.Width * scale;
            float displayHeight = image.Height * scale;

            ImageFragment imageFragment = widget.AddImageFragment(image, 2, (boxHeight - displayHeight) / 2);
            imageFragment.ScaleX = displayWidth;
            imageFragment.ScaleY = displayHeight;
        }
        void BuildSignatureAppearance(SignatureWidgetAnnotation widget, DXImage image, SignatureOptions signatureOptions, AppearanceSettings appearanceSettings) {
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
            if(appearanceSettings.ShowDate) {
                widget.AddTextFragment($"Date: {signatureOptions.SigningTime?.ToString("yyyy'.'MM'.'dd HH':'mm':'ss zzz")}", textX, textY).FontSize = fontSize;
            }
        }
        class AppearanceSettings {
            public bool ShowDate { get; set; } = false;
            public bool ShowLocation { get; set; } = false;
            public bool ShowName { get; set; } = false;
            public bool ShowReason { get; set; } = false;
        }
        #endregion
    }
}
