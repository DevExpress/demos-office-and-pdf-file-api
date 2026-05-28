using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.DocumentElements;
#region Code
using DevExpress.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.DocumentElements {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfAttachmentsController : PdfControllerBase {
        [HttpPost("pdf-attachments/attach")]
        public async Task<IActionResult> AttachFileToPdf(IFormFile file, IFormFile attachment, [FromForm] PdfAttachmentsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            using var attachmentStream = new MemoryStream();
            await attachment.CopyToAsync(attachmentStream);
            attachmentStream.Position = 0;

            var outputStream = AttachFile(documentStream, attachment.FileName, "DevExpress Logo Image", GetMimeContentType("png"), attachmentStream.ToArray(), model.CreateAnnotation);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AttachFile(Stream documentStream, string fileName, string fileDescription, string fileMimeType, byte[] fileData, bool createAnnotation) {
            using var processor = new PdfDocumentProcessor();
            
            processor.LoadDocument(documentStream);

            var attachment = new PdfFileAttachment() {
                CreationDate = DateTime.UtcNow,
                FileName = fileName,
                Data = fileData,
                MimeType = fileMimeType,
                Description = fileDescription,
                ModificationDate = DateTime.UtcNow,
                Relationship = PdfAssociatedFileRelationship.Supplement
            };

            if(createAnnotation) {
                var page = processor.Document.Pages[0];
                var pageFacade = processor.DocumentFacade.Pages[0];
                pageFacade.AddFileAttachmentAnnotation(new PdfPoint(page.MediaBox.Right - 50, page.MediaBox.Bottom + 50), attachment, PdfFileAttachmentAnnotationIconName.PaperClip);
            } else
                processor.AttachFile(attachment);

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            return outputStream;
        }
        #endregion
    }
}
