using DevExpress.Pdf;

namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.Forms {
    public class PdfImportExportFormDataModel : PdfModelBase {
        public PdfFormDataFormat FormDataExportFormat { get; set; } = PdfFormDataFormat.Xml;
        public PdfFormDataFormat FormDataImportFormat { get; set; } = PdfFormDataFormat.Xml;
        internal List<FormatInfo> FormDataFormats { get; } = new() {
            new() { Text = "XML", Value = PdfFormDataFormat.Xml },
            new() { Text = "FDF (Forms Data Format)", Value = PdfFormDataFormat.Fdf },
            new() { Text = "XFDF (XML Forms Data Format)", Value = PdfFormDataFormat.Xfdf },
            new() { Text = "Plain Text", Value = PdfFormDataFormat.Txt }
        };
        internal readonly List<string> FormDataExtensions = new() {
            ".xml"
        };
    }
}
