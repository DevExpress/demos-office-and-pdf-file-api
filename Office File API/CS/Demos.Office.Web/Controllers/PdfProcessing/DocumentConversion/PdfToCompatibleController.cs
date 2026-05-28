using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.DocumentConversion;
#region Code
using DevExpress.Pdf;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.DocumentConversion {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfToCompatibleController : PdfControllerBase {
        static Stream conversionReportStream;

        [HttpPost("document-conversion/pdf-to-compatible")]
        public async Task<IActionResult> ConvertPdfToCompatible(
            IFormFile file,
            [FromForm] PdfToCompatibleModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStreams = ConvertToCompatible(documentStream, model.CompatibilityLevel);
            outputStreams[0].Seek(0, SeekOrigin.Begin);
            outputStreams[1].Seek(0, SeekOrigin.Begin);

            conversionReportStream = new MemoryStream();
            outputStreams[1].CopyTo(conversionReportStream);
            conversionReportStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStreams[0], model.OutputFormat, model.OutputName);
        }
        [HttpPost("document-conversion/pdf-to-compatible-report")]
        public async Task<IActionResult> GetPdfConversionReport() {
            return CreateOutputFile(conversionReportStream, "txt", "Conversion Report");
        }
        #region Code
        IReadOnlyList<Stream> ConvertToCompatible(Stream documentStream, PdfCompatibility compatibilityLevel) {
            using var converter = new PdfDocumentConverter(documentStream);

            var outputStream = new MemoryStream();
            var reportStream = new MemoryStream();
            var streamWriter = new StreamWriter(reportStream);

            converter.Convert(compatibilityLevel);
            converter.SaveDocument(outputStream);

            streamWriter.WriteLine($"Conversion Status : {converter.ConversionReport.ConversionStatus}");
            if(converter.ConversionReport.Issues.Any()) {
                streamWriter.WriteLine("Found Issues :");
                streamWriter.Write(string.Join("\r\n", converter.ConversionReport.Issues.Select(x => $"{x.Severity} : {x.Message}")));
            }

            streamWriter.Flush();

            reportStream.Position = 0;
            outputStream.Position = 0;
            
            return new List<Stream> { outputStream, reportStream };
        }
        #endregion Code
    }
}
