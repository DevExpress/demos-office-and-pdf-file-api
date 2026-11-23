using DevExpress.Docs.Word;

namespace BlazorDemo.OfficeFileApi.Services {
    public static class NewWordFormatHelper {
        public static DocumentFormat GetDocumentFormat(string format) {
            return format.ToLowerInvariant() switch {
                "doc" => DocumentFormat.Doc,
                "docx" => DocumentFormat.Docx,
                "docm" => DocumentFormat.Docm,
                "dot" => DocumentFormat.Dot,
                "dotx" => DocumentFormat.Dotx,
                "dotm" => DocumentFormat.Dotm,
                "rtf" => DocumentFormat.Rtf,
                "html" => DocumentFormat.Html,
                "mht" => DocumentFormat.Mht,
                "txt" => DocumentFormat.PlainText,
                "odt" => DocumentFormat.Odt,
                "xml" => DocumentFormat.WordML,
                _ => DocumentFormat.Docx
            };
        }
    }
}