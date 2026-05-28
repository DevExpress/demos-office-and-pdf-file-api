using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.ManagePages;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.ManagePages {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfManagePagesController : PdfControllerBase {
        [HttpPost("manage-pages/add")]
        public async Task<IActionResult> AddPageToPdf(IFormFile file, [FromForm] PdfManagePagesModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var (width, height) = NewPageSizeHelper.GetDimensions(model.NewPageSize);
            var outputStream = AddPage(documentStream, width, height);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AddPage(Stream documentStream, double width, double height) {
            using var doc = new PdfDocument(documentStream);

            doc.Pages.Add((float)width, (float)height);

            var outputStream = new MemoryStream();
            doc.Save(outputStream);

            return outputStream;
        }
        #endregion

        [HttpPost("manage-pages/delete")]
        public async Task<IActionResult> DeletePagesFromPdf(IFormFile file, [FromForm] PdfManagePagesModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var pageRange = RangeHelper.GetIndexesByRange(model.ExportMode, model.PageRange);

            var outputStream = DeletePages(documentStream, pageRange);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream DeletePages(Stream documentStream, IEnumerable<int> pageRange) {
            using var doc = new PdfDocument(documentStream);

            foreach(var index in pageRange) {
                if(index < 0 || index >= doc.Pages.Count)
                    continue;

                doc.Pages.RemoveAt(index);
            }

            var outputStream = new MemoryStream();
            doc.Save(outputStream);

            return outputStream;
        }
        #endregion

        [HttpPost("manage-pages/duplicate")]
        public async Task<IActionResult> DuplicatePagesInPdf(IFormFile file, [FromForm] PdfManagePagesModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var pageRange = RangeHelper.GetIndexesByRange(model.ExportMode, model.PageRange);

            var outputStream = DuplicatePages(documentStream, pageRange);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream DuplicatePages(Stream documentStream, IEnumerable<int> pageRange) {
            using var doc = new PdfDocument(documentStream);

            foreach(var index in pageRange) {
                if(index < 0 || index >= doc.Pages.Count)
                    continue;

                doc.Pages.Add(doc.Pages[index].Clone());
            }

            var outputStream = new MemoryStream();
            doc.Save(outputStream);

            return outputStream;
        }
        #endregion
    }
}
