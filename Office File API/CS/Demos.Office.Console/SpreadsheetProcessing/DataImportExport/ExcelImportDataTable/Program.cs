using DevExpress.Drawing;
using DevExpress.Spreadsheet;
using System.Data;
using System.Drawing;

class Program {
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Imported_Data.xlsx";

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var outputStream = ImportDataTable(CreateEmployeesTable(), DocumentFormat.Xlsx);
            using var fileStream = File.Create(outputPath);
            outputStream.CopyTo(fileStream);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static Stream ImportDataTable(DataTable sourceTable, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();

        var firstWorksheet = workbook.Worksheets.First();
        firstWorksheet.Name = "Employees";
        firstWorksheet["A1"].Value = "Import data from an Employees data table:";

        var formatting = firstWorksheet["B1:G1"].BeginUpdateFormatting();
        formatting.Font.Bold = true;
        firstWorksheet["B1:G1"].EndUpdateFormatting(formatting);

        var importOptions = new DataImportOptions() {
            Converter = new DataValueConverter()
        };

        firstWorksheet.Import(sourceTable, true, 0, 1, importOptions);
        firstWorksheet.Columns.AutoFit(0, 6);

        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }

    static DataTable CreateEmployeesTable() {
        var employees = new[] {
            new EmployeeRecord(1, "Nancy Davolio", "Sales Representative", "Seattle", CreateSampleImageBase64(Color.LightBlue)),
            new EmployeeRecord(2, "Andrew Fuller", "Vice President, Sales", "Tacoma", CreateSampleImageBase64(Color.LightGreen)),
            new EmployeeRecord(3, "Janet Leverling", "Sales Representative", "Kirkland", CreateSampleImageBase64(Color.LightSalmon))
        };

        DataTable employeesDataTable = new DataTable();
        employeesDataTable.Columns.Add("EmployeeID");
        employeesDataTable.Columns.Add("FullName");
        employeesDataTable.Columns.Add("Title");
        employeesDataTable.Columns.Add("City");
        employeesDataTable.Columns.Add("Picture");

        foreach(var employee in employees) {
            employeesDataTable.Rows.Add(employee.EmployeeID, employee.FullName, employee.Title, employee.City, employee.PictureBase64);
        }

        return employeesDataTable;
    }

    static string CreateSampleImageBase64(Color color) {
        using var bitmap = new DXBitmap(48, 48);
        using(DXGraphics graphics = DXGraphics.FromImage(bitmap)) {
            graphics.Clear(color);
            graphics.FillEllipse(DXBrushes.White, new RectangleF(8, 8, 32, 32));
        }

        using var stream = new MemoryStream();
        bitmap.Save(stream, DXImageFormat.Png);
        return Convert.ToBase64String(stream.ToArray());
    }

    sealed class DataValueConverter : IDataValueConverter {
        public bool TryConvert(object value, int index, out CellValue result) {
            if(value is string strValue) {
                try {
                    result = DXImage.FromBase64String(strValue);
                }
                catch {
                    result = Int32.TryParse(strValue, out int intValue) ? intValue : strValue;
                }
                return true;
            }

            result = CellValue.TryCreateFromObject(value);
            return true;
        }
    }

    sealed record EmployeeRecord(int EmployeeID, string FullName, string Title, string City, string PictureBase64);
}
