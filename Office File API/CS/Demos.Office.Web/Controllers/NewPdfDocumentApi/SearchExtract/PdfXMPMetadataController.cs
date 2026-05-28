using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.SearchExtract;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.SearchExtract {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfXMPMetadataController : PdfControllerBase {
        [HttpPost("pdf-search-extract/extract-metadata")]
        public async Task<IActionResult> GetPdfXmpMetadata(IFormFile file, [FromForm] PdfXMPMetadataModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = GetXmpMetadata(documentStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, "txt", "Metadata");
        }

        #region Code
        Stream GetXmpMetadata(Stream documentStream) {
            using var doc = new PdfDocument(documentStream);

            var outputStream = new MemoryStream();
            var writer = new StreamWriter(outputStream);
            if (doc.Metadata.Xmp != null)
                writer.Write(doc.Metadata.Xmp.ToXmlString());
            writer.Flush();

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion

        [HttpPost("pdf-search-extract/remove-metadata")]
        public async Task<IActionResult> RemovePdfXmpMetadata(IFormFile file, [FromForm] PdfXMPMetadataModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = RemoveXmpMetadata(documentStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream RemoveXmpMetadata(Stream documentStream) {
            using var doc = new PdfDocument(documentStream);

            doc.Metadata.Xmp = null;

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion

        [HttpPost("pdf-search-extract/add-metadata")]
        public async Task<IActionResult> AddPdfXmpMetadata(
            IFormFile file,
            IFormFile metadata,
            [FromForm] PdfXMPMetadataModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            using var metadataStream = new MemoryStream();
            await metadata.CopyToAsync(metadataStream);
            metadataStream.Position = 0;

            var outputStream = AddXmpMetadata(documentStream, metadataStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AddXmpMetadata(Stream documentStream, Stream metadataStream) {
            using var doc = new PdfDocument(documentStream);

            doc.Metadata.Xmp = XmpMetadata.FromStream( metadataStream);

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion
    }
}
