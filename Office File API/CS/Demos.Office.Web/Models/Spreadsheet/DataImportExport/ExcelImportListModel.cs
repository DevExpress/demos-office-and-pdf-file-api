using DevExpress.Spreadsheet;
using System.ComponentModel.DataAnnotations;
using DevExpress.Data.Utils.Security;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DataImportExport {
    public class ExcelImportListModel : ExcelFileModelBase {
        // Override OutputFormat validation to include Excel formats
        [RegularExpression(@"^(xlsx|xls|xlsm|xltx|xlt|xltm|xlsb)$",
            ErrorMessage = "Invalid output format. Supported formats: xlsx, xls, xlsm, xltx, xlt, xltm, xlsb")]

        public override string OutputFormat { get; set; } = "xlsx";
        public override string OutputName { get; set; } = "Imported_Users";

        // Not used; added to bind DxListBox @bind-Value
        internal UserModel SelectedUser { get; set; } = new UserModel();

        public IList<UserModel> Users { get; set; } = new InMemoryNWindDataService().Employees
            .Take(3)
            .Select(
            employee => new UserModel() {
                FirstName = employee.FirstName,
                LastName = employee.LastName,
                BirthDate = new DateTime(StrongRandom.Next(1975, 2000), StrongRandom.Next(1, 12), StrongRandom.Next(1, 28)),
            }
        ).ToList();

        public IEnumerable<string> SelectedFields { get; set; } = new List<string>{
            "FirstName",
            "LastName",
            "BirthDate"
        };
        internal readonly List<object> FieldNames = new() {
             new { Text = "First Name", Value = "FirstName" },
             new { Text = "Last Name", Value = "LastName" },
             new { Text = "Birth Date", Value = "BirthDate" }
        };
        public override List<FormatInfo> OutputFormats { get; set; } = new() {
           new FormatInfo() { Text = "Excel Workbook (*.xlsx)", Value = "xlsx" },
           new FormatInfo() { Text = "Excel 97-2003 Workbook (*.xls)", Value = "xls" },
           new FormatInfo() { Text = "Excel Macro-Enabled Workbook (*.xlsm)", Value = "xlsm" },
           new FormatInfo() { Text = "Excel Template (*.xltx)", Value = "xltx" },
           new FormatInfo() { Text = "Excel 97-2003 Template (*.xlt)", Value = "xlt" },
           new FormatInfo() { Text = "Excel Macro-Enabled Template (*.xltm)", Value = "xltm" },
           new FormatInfo() { Text = "Excel Binary Workbook (*.xlsb)", Value = "xlsb" },
        };
    }
    public class UserModel {
        [Required(ErrorMessage = "First Name is required")]
        [MinLength(1, ErrorMessage = "First Name can't be empty")]
        [MaxLength(30, ErrorMessage = "First Name too long")]
        public string FirstName { get; set; } = string.Empty;

        [Required(ErrorMessage = "Last Name is required")]
        [MinLength(1, ErrorMessage = "Last Name can't be empty")]
        [MaxLength(30, ErrorMessage = "Last Name too long")]
        public string LastName { get; set; } = string.Empty;

        DateTime? _birthDate;
        [Required]
        public DateTime? BirthDate {
            get { return _birthDate; }
            set { _birthDate = value?.Date; }
        }
    }
}
