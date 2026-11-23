using BlazorDemo.OfficeFileApi.Models.PdfProcessing;
using DevExpress.Docs.Pdf;
using DevExpress.Office.DigitalSignatures;
using DevExpress.Pdf;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentProtection {
    public class PdfDigitalSignatureModel : PdfModelBase {
        public PdfSignatureAppearanceType AppearanceType { get; set; } = PdfSignatureAppearanceType.None;
        public string DateTimeFormat { get; set; } = "MM/dd/yyyy hh:mm";
        public List<object> HashAlgorithms { get; } = new() {
            new { Text = "SHA256", Value = HashAlgorithm.SHA256},
            new { Text = "SHA384", Value = HashAlgorithm.SHA384},
            new { Text = "SHA512", Value = HashAlgorithm.SHA512}
        };
        public bool ShowDate { get; set; } = true;
        public bool ShowLocation { get; set; } = true;
        public bool ShowName { get; set; } = true;
        public bool ShowReason { get; set; } = true;
        public HashAlgorithmType TsaHashAlgorithm { get; set; } = HashAlgorithmType.SHA256;
        public HashAlgorithm CertificateHashAlgorithm { get; set; } = HashAlgorithm.SHA256;
        public bool UseCustomAppearance { get; set; }
    }
}
