using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.SearchExtract;
#region Code
using DevExpress.Pdf;
using DevExpress.Pdf.Xmp;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.SearchExtract {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfXMPMetadataController : PdfControllerBase {
        [HttpPost("pdf-search-extract/extract-metadata")]
        public async Task<IActionResult> GetPdfXMPMetadata(IFormFile file, [FromForm] PdfXMPMetadataModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = GetMetadata(documentStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream GetMetadata(Stream documentStream) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            var xmpDoc = XmpDocument.FromString(processor.Document.Metadata.Data);

            var outputStream = new MemoryStream();
            xmpDoc.Serialize(outputStream, new XmpSerializationOptions());

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion

        [HttpPost("pdf-search-extract/remove-metadata")]
        public async Task<IActionResult> RemovePdfXMPMetadata(IFormFile file, [FromForm] PdfXMPMetadataModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = RemoveMetadata(documentStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream RemoveMetadata(Stream documentStream) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            var xmpDoc = XmpDocument.FromString(processor.Document.Metadata.Data);
            xmpDoc.Clear();
            processor.Document.SetMetadata(xmpDoc);

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion

        [HttpPost("pdf-search-extract/add-metadata")]
        public async Task<IActionResult> AddPdfXMPMetadata(
            IFormFile file,
            IFormFile metadata,
            [FromForm] PdfXMPMetadataModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            using var metadataStream = new MemoryStream();
            await metadata.CopyToAsync(metadataStream);
            metadataStream.Position = 0;

            var outputStream = AddMetadata(documentStream, metadataStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AddMetadata(Stream documentStream, Stream metadataStream) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            var metadata = XmpDocument.FromStream(metadataStream);
            processor.Document.SetMetadata(metadata);

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion
    }
}
