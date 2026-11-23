using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentConversion;
#region Code
using DevExpress.Docs.Pdf;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.DocumentConversion {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfConvertToConformanceController : PdfControllerBase {
        static Stream conversionReportStream;

        [HttpPost("pdf-document-conversion/convert-to-conformance")]
        public async Task<IActionResult> ConvertToConformance(
            IFormFile file,
            [FromForm] PdfConvertToConformanceModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStreams = ConvertToPdfA(documentStream, model.Compatibility);
            outputStreams[0].Seek(0, SeekOrigin.Begin);
            outputStreams[1].Seek(0, SeekOrigin.Begin);

            conversionReportStream = new MemoryStream();
            outputStreams[1].CopyTo(conversionReportStream);
            conversionReportStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStreams[0], model.OutputFormat, model.OutputName);
        }

        [HttpPost("pdf-document-conversion/convert-to-conformance-report")]
        public async Task<IActionResult> GetPdfConversionReport() {
            return CreateOutputFile(conversionReportStream, "txt", "Conversion Report");
        }

        #region Code
        IReadOnlyList<Stream> ConvertToPdfA(Stream documentStream, PdfACompatibility compatibility) {
            using var document = new PdfDocument(documentStream);

            var outputStream = new MemoryStream();
            var reportStream = new MemoryStream();
            var streamWriter = new StreamWriter(reportStream);

            var result = document.ConvertToPdfA(compatibility);
            document.Save(outputStream);

            streamWriter.WriteLine($"Conversion Status : {result.Status}");
            if(result.Issues.Any()) {
                streamWriter.WriteLine("Found Issues :");
                streamWriter.Write(string.Join("\r\n", result.Issues.Select(x => $"{x.Severity} : {x.Message}")));
            }

            streamWriter.Flush();

            reportStream.Position = 0;
            outputStream.Position = 0;

            return new List<Stream> { outputStream, reportStream };
        }
        #endregion Code
    }
}
