using System.ComponentModel.DataAnnotations;
using System.Data;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DataImportExport {
    public class ExcelImportDataTableModel : ExcelFileModelBase {
        // Override OutputFormat validation to include Excel formats
        [RegularExpression(@"^(xlsx|xls|xlsm|xltx|xlt|xltm|xlsb)$",
            ErrorMessage = "Invalid output format. Supported formats: xlsx, xls, xlsm, xltx, xlt, xltm, xlsb")]

        public override string OutputFormat { get; set; } = "xlsx";
        public override string OutputName { get; set; } = "Imported_Data";

        public DataTable EmployeesTable {
            get {
                var employees = new InMemoryNWindDataService().Employees;

                DataTable employeesDataTable = new DataTable();  
                employeesDataTable.Columns.Add("EmployeeID");
                employeesDataTable.Columns.Add("FullName");
                employeesDataTable.Columns.Add("Title");
                employeesDataTable.Columns.Add("City");
                employeesDataTable.Columns.Add("Picture");

                foreach (var employee in employees) {
                    employeesDataTable.Rows.Add(
                        employee.EmployeeID,
                        employee.FullName,
                        employee.Title,
                        employee.City,
                        Convert.ToBase64String(employee.Picture ?? Array.Empty<byte>()));
                }
                return employeesDataTable;
            }
        }  

        public override List<FormatInfo> OutputFormats { get; set; } = new() {
           new FormatInfo { Text = "Excel Workbook (*.xlsx)", Value = "xlsx" },
           new FormatInfo { Text = "Excel 97-2003 Workbook (*.xls)", Value = "xls" },
           new FormatInfo { Text = "Excel Macro-Enabled Workbook (*.xlsm)", Value = "xlsm" },
           new FormatInfo { Text = "Excel Template (*.xltx)", Value = "xltx" },
           new FormatInfo { Text = "Excel 97-2003 Template (*.xlt)", Value = "xlt" },
           new FormatInfo { Text = "Excel Macro-Enabled Template (*.xltm)", Value = "xltm" },
           new FormatInfo { Text = "Excel Binary Workbook (*.xlsb)", Value = "xlsb" },
        };
    }
}
