using DevExpress.Docs.Pdf;
using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentProtection {
    public class PdfEncryptPasswordModel : PdfModelBase {
        public DocumentDataExtractionPermissions DocumentDataExtractionPermissions { get; set; } = DocumentDataExtractionPermissions.Allowed;
        public List<object> DocumentDataExtractionPermissionsCollection { get; } = new() {
            new { Text = "Not Allowed", Value = DocumentDataExtractionPermissions.NotAllowed },
            new { Text = "Only Accessibility", Value = DocumentDataExtractionPermissions.Accessibility },
            new { Text = "Allowed", Value = DocumentDataExtractionPermissions.Allowed }
        };
        public DocumentInteractivityPermissions DocumentInteractivityPermissions { get; set; } = DocumentInteractivityPermissions.Allowed;
        public List<object> DocumentInteractivityPermissionsCollection { get; } = new() {
            new { Text = "Not Allowed", Value = DocumentInteractivityPermissions.NotAllowed },
            new { Text = "Only Form Filling and Signing", Value = DocumentInteractivityPermissions.FormFillAndSign },
            new { Text = "Allowed", Value = DocumentInteractivityPermissions.Allowed }
        };
        public DocumentModificationPermissions DocumentModificationPermissions { get; set; } = DocumentModificationPermissions.Allowed;
        public List<object> DocumentModificationPermissionsCollection { get; } = new() {
            new { Text = "Not Allowed", Value = DocumentModificationPermissions.NotAllowed },
            new { Text = "Only Document Assembling", Value = DocumentModificationPermissions.DocumentAssembling },
            new { Text = "Allowed", Value = DocumentModificationPermissions.Allowed }
        };
        public DocumentPrintPermissions DocumentPrintingPermissions { get; set; } = DocumentPrintPermissions.Allowed;
        public List<object> DocumentPrintingPermissionsCollection { get; } = new() {
            new { Text = "Not Allowed", Value = DocumentPrintPermissions.NotAllowed },
            new { Text = "Only Low Quality", Value = DocumentPrintPermissions.LowQuality },
            new { Text = "Allowed", Value = DocumentPrintPermissions.Allowed },
        };
        [Required(ErrorMessage = "An encryption algorithm is required")]
        public EncryptionAlgorithm EncryptionAlgorithm { get; set; } = EncryptionAlgorithm.AES256;
        public List<object> EncryptionAlgorithms { get; } = new() {
            new { Text = "AES128", Value = EncryptionAlgorithm.AES128 },
            new { Text = "AES256", Value = EncryptionAlgorithm.AES256 }
        };
        [Required(ErrorMessage = "Password is required")]
        [MinLength(1, ErrorMessage = "Password cannot be empty")]
        public string OwnerPassword { get; set; }
        [MinLength(1, ErrorMessage = "Password cannot be empty")]
        public string UserPassword { get; set; }
    }
}
