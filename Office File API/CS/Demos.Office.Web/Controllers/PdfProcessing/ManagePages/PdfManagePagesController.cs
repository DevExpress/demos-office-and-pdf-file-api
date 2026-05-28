using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.ManagePages;
#region Code
using DevExpress.Pdf;
using DevExpress.Drawing;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.ManagePages {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfManagePagesController : PdfControllerBase {
        [HttpPost("manage-pages/add")]
        public async Task<IActionResult> AddPageToPdf(IFormFile file, [FromForm] PdfManagePagesModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = AddPage(documentStream, new PdfRectangle(0, 0, model.Width, model.Height));
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AddPage(Stream documentStream, PdfRectangle pageSize) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            var newPage = processor.AddNewPage(pageSize);

            var text = "New Page";
            var font = new DXFont("Arial", 25);
            using var brush = new DXSolidBrush(System.Drawing.Color.Black);
            using var g = processor.CreateGraphicsPageSystem();                

            var textSize = g.MeasureString(text, font);

            g.DrawString(text, font, brush, (float)(pageSize.Width - textSize.Width) / 2, (float)(pageSize.Height - textSize.Height) / 2);
            g.AddToPageForeground(newPage);

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            return outputStream;
        }
        #endregion

        [HttpPost("manage-pages/delete")]
        public async Task<IActionResult> DeletePagesFromPdf(IFormFile file, [FromForm] PdfManagePagesModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var pageRange = RangeHelper.GetIndexesByRange(model.ExportMode, model.PageRange).Select(x => x + 1).Reverse();

            var outputStream = DeletePages(documentStream, pageRange);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream DeletePages(Stream documentStream, IEnumerable<int> pageRange) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            foreach(var index in pageRange) {
                if(index < 1 || index > processor.Document.Pages.Count)
                    continue;

                processor.DeletePage(index);
            }

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            return outputStream;
        }
        #endregion

        [HttpPost("manage-pages/duplicate")]
        public async Task<IActionResult> DuplicatePagesInPdf(IFormFile file, [FromForm] PdfManagePagesModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var pageRange = RangeHelper.GetIndexesByRange(model.ExportMode, model.PageRange);
            pageRange.Reverse();

            var outputStream = DuplicatePages(documentStream, pageRange);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream DuplicatePages(Stream documentStream, IEnumerable<int> pageRange) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            foreach(var index in pageRange) {
                if(index < 0 || index >= processor.Document.Pages.Count)
                    continue;

                processor.Document.Pages.Add(processor.Document.Pages[index]);
            }

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            return outputStream;
        }
        #endregion
    }
}
