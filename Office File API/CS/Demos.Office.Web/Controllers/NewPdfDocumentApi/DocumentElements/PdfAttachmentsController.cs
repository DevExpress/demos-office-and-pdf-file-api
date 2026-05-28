using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentElements;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.DocumentElements {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfAttachmentsController : PdfControllerBase {
        [HttpPost("pdf-attachments/attach")]
        public async Task<IActionResult> AttachFileToPdf(IFormFile file, IFormFile attachment, [FromForm] PdfAttachmentsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            using var attachmentStream = new MemoryStream();
            await attachment.CopyToAsync(attachmentStream);
            attachmentStream.Position = 0;

            var outputStream = AttachFile(documentStream, attachment.FileName, "DevExpress Logo Image", GetMimeContentType("png"), attachmentStream.ToArray());
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AttachFile(Stream documentStream, string fileName, string fileDescription, string fileMimeType,
            byte[] fileData) {
            using var doc = new PdfDocument(documentStream);
            var attachment = new Attachment() {
                CreationDate = DateTime.UtcNow,
                FileName = fileName,
                Data = fileData,
                MimeType = fileMimeType,
                Description = fileDescription,
                ModificationDate = DateTime.UtcNow,
                Relationship = AssociatedFileRelationship.Supplement
            };
            doc.Attachments.Add(attachment);
            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            return outputStream;
        }

        #endregion
    }
}
