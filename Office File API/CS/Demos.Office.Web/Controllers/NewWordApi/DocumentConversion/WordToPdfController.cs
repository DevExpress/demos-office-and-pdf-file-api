#region Code

using DevExpress.Docs.Pdf;
using DevExpress.Docs.Word;
using DevExpress.Docs.Word.Export;
        #endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentConversion;

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.DocumentConversion {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordToPdfController : WordControllerBase {
        [HttpPost("word-to-pdf/convert")]
        public async Task<IActionResult> ConvertWordToPdf(IFormFile file, [FromForm] WordToPdfConversionModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var pdfStream = ConvertToPdf(inputStream, model.PdfA, model.PdfUA,
                model.Author ?? string.Empty, model.Title ?? string.Empty, model.Subject ?? string.Empty, model.Keywords ?? string.Empty, model.ExportControlsAsAcroFormFields);

            pdfStream.Position = 0;

            return CreateOutputFile(pdfStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ConvertToPdf(Stream inputStream, PdfACompatibility pdfAConformance, PdfUACompatibility pdfUACompatibility,
                            string author, string title, string subject, string keywords, bool exportControlsAsAcroForm) {
            // Load the input document
            using var document = new WordDocument(inputStream);

            // Create output stream
            var outputStream = new MemoryStream();

            // Export to PDF with specified settings
            var exportOptions = new PdfExportOptions() {
                PdfACompatibility = pdfAConformance,
                PdfUACompatibility = pdfUACompatibility,
                ExportInteractiveControls = exportControlsAsAcroForm
            };
            exportOptions.DocumentOptions.Author = author;
            exportOptions.DocumentOptions.Title = title;
            exportOptions.DocumentOptions.Subject = subject;
            exportOptions.DocumentOptions.Keywords = keywords;

            document.ExportToPdf(outputStream, exportOptions);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}