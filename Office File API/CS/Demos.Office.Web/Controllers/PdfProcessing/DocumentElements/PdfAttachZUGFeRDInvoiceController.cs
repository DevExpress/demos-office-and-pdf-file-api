using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.DocumentElements;
#region Code
using DevExpress.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.DocumentElements {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfAttachZUGFeRDInvoiceController : PdfControllerBase {
        [HttpPost("document-elements/pdf-attach-zugferd")]
        public async Task<IActionResult> ConvertPdfAttachZUGFeRDInvoice(IFormFile file, IFormFile invoice, [FromForm] PdfAttachZUGFeRDInvoiceModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            using var invoiceStream = new MemoryStream();
            await invoice.CopyToAsync(invoiceStream);
            invoiceStream.Position = 0;

            var outputStream = AttachZugferd(documentStream, invoiceStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AttachZugferd(Stream documentStream, Stream invoiceStream) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            processor.Document.AttachZugferdInvoice(invoiceStream);

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            return outputStream;
        }
        #endregion
    }
}
