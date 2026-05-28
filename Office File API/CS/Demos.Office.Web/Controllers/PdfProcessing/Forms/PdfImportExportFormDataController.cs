using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.Forms;
#region Code
using DevExpress.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.Forms {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfImportExportFormDataController : PdfControllerBase {
        [HttpPost("pdf-form/import")]
        public async Task<IActionResult> ImportPdfFormData(IFormFile file, IFormFile formData, [FromForm] PdfImportExportFormDataModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            using var formDataStream = new MemoryStream();
            await formData.CopyToAsync(formDataStream);
            formDataStream.Position = 0;

            var outputStream = ImportFormData(documentStream, formDataStream, model.FormDataImportFormat);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ImportFormData(Stream documentStream, Stream formDataStream, PdfFormDataFormat formDataFormat) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            processor.ApplyFormData(new PdfFormData(formDataStream, formDataFormat));

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion

        [HttpPost("pdf-form/export")]
        public async Task<IActionResult> ExportPdfFormData(IFormFile file, [FromForm] PdfImportExportFormDataModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = ExportFormData(documentStream, model.FormDataExportFormat);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.FormDataExportFormat.ToString().ToLowerInvariant(), "Data");
        }

        #region Code
        Stream ExportFormData(Stream documentStream, PdfFormDataFormat formDataFormat) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            var formData = processor.GetFormData();

            var outputStream = new MemoryStream();
            formData.Save(outputStream, formDataFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion
    }
}
