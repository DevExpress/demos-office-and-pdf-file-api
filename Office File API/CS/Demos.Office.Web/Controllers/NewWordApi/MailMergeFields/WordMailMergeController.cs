#region Code

using DevExpress.Docs.Word;
using DevExpress.Drawing;
        #endregion Code

using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.MailMergeFields;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.MailMergeFields {
    [ApiController]
    [Route("api/newwordapi/mail-merge")]
    public class WordMailMergeController : WordControllerBase {
        [HttpPost("process")]
        public async Task<IActionResult> CreateMailMergedDocument(IFormFile file, [FromForm] WordMailMergeModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var selectedEmployee = model.Employees.FirstOrDefault(employee => employee.EmployeeID == model.EmployeeId);
            if(selectedEmployee != null) {
                var selectedCustomer = model.Customers.FirstOrDefault(customer => customer.CustomerID == model.CustomerId);
                if(selectedCustomer != null) {
                    selectedEmployee.Customer = selectedCustomer;
                }
            }

            // Process track changes on the template
            var outputStream = MailMergeDocument(inputStream, selectedEmployee, model.IsPdfOutput, model.OutputFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        const string EmployeePhotoPrefix = "dbimg://";

        MemoryStream MailMergeDocument(Stream stream, Employee selectedEmployee, bool isPdfOutput, string outputFormat) {
            using var document = new WordDocument(stream);

            if(selectedEmployee == null)
                return null;

            ReplaceEmployeePhotoFields(document.Body, selectedEmployee.EmployeeID);

            var options = new MailMergeOptions {
                DataSource = new List<Employee>() { selectedEmployee },
                FirstRecordIndex = 0,
                LastRecordIndex = 0
            };

            using var mergedDocument = document.MailMerge(options);

            var outputStream = new MemoryStream();
            if(isPdfOutput)
                mergedDocument.ExportToPdf(outputStream);
            else
                mergedDocument.Save(outputStream, NewWordFormatHelper.GetDocumentFormat(outputFormat));

            outputStream.Position = 0;
            return outputStream;
        }

        void ReplaceEmployeePhotoFields(IEnumerable<Block> blocks, int employeeId) {
            foreach(var block in blocks) {
                if(block is Paragraph paragraph)
                    ReplaceEmployeePhotoFields(paragraph, employeeId);
                else if(block is Table table) {
                    foreach(var row in table.Rows) {
                        foreach(var cell in row.Cells)
                            ReplaceEmployeePhotoFields(cell.Content, employeeId);
                    }
                }
            }
        }

        void ReplaceEmployeePhotoFields(Paragraph paragraph, int employeeId) {
            var inlineElements = paragraph.InlineElements;
            for(int i = 0; i < inlineElements.Count; i++) {
                if(inlineElements[i] is Field field && IsEmployeePhotoField(field))
                    inlineElements[i] = CreateEmployeePhoto(employeeId);
            }
        }

        bool IsEmployeePhotoField(Field field) {
            var fieldCode = field.Code.GetText().TrimStart();
            return fieldCode.StartsWith("INCLUDEPICTURE", StringComparison.OrdinalIgnoreCase)
                   && fieldCode.Contains(EmployeePhotoPrefix, StringComparison.OrdinalIgnoreCase);
        }

        PictureShape CreateEmployeePhoto(int employeeId) {
            var fileName = $"Documents/Photo{employeeId}.jpeg";
            var imageStream = new MemoryStream(System.IO.File.ReadAllBytes(fileName));
            var result = new PictureShape(DXImage.FromStream(imageStream));
            return result;
        }
        #endregion Code
    }
}