using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using System.Data;
using System.Text.Json;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "TopTradingPartners.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Exported_Data.json";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            DataTable exportedTable = ExportDataTable(inputStream);
            var json = JsonSerializer.Serialize(ConvertToDTO(exportedTable), new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static DataTable ExportDataTable(Stream inputStream) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);

        Worksheet worksheet = workbook.Worksheets[0];
        CellRange range = worksheet.Tables[0].Range;

        DataTable dataTable = worksheet.CreateDataTable(range, true);
        dataTable.Columns["As Of"].DataType = typeof(string);

        DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
        exporter.CellValueConversionError += Exporter_CellValueConversionError;

        EmptyCellValueConverter myconverter = new EmptyCellValueConverter();
        exporter.Options.CustomConverters.Add("As Of", myconverter);

        myconverter.EmptyCellValue = "N/A";
        exporter.Options.ConvertEmptyCells = true;
        exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = false;
        exporter.Export();

        return dataTable;
    }

    static void Exporter_CellValueConversionError(object sender, CellValueConversionErrorEventArgs e) {
        e.DataTableValue = null;
        e.Action = DataTableExporterAction.Continue;
    }

    class EmptyCellValueConverter : ICellValueToColumnTypeConverter {
        public bool SkipErrorValues { get; set; }
        public CellValue EmptyCellValue { get; set; }

        public ConversionResult Convert(Cell readOnlyCell, CellValue cellValue, Type dataColumnType, out object result) {
            result = DBNull.Value;

            ConversionResult converted = ConversionResult.Success;
            if(cellValue.IsEmpty) {
                result = EmptyCellValue;
                return converted;
            }
            if(cellValue.IsError) {
                result = "N/A";
                return ConversionResult.Success;
            }
            result = String.Format("{0:MMMM-yyyy}", cellValue.DateTimeValue);
            return converted;
        }
    }

    static List<TradingPartnerDTO> ConvertToDTO(DataTable exportedTable) {
        List<TradingPartnerDTO> dtos = new List<TradingPartnerDTO>();
        foreach(DataRow row in exportedTable.Rows) {
            dtos.Add(new TradingPartnerDTO() {
                Country = (string)row["Country"],
                Exports = (double)row["Exports"],
                Imports = (double)row["Imports"],
                Balance = (double)row["Balance"],
                Exports1YChg = (double)row["Exports 1Y Chg"],
                Imports1YChg = (double)row["Imports 1Y Chg"],
                Balance1YChg = (double)row["Balance 1Y Chg"],
                AsOf = (string)row["As Of"],
            });
        }
        return dtos;
    }

    class TradingPartnerDTO {
        public string Country { get; set; }
        public double Exports { get; set; }
        public double Imports { get; set; }
        public double Balance { get; set; }
        public double Exports1YChg { get; set; }
        public double Imports1YChg { get; set; }
        public double Balance1YChg { get; set; }
        public string AsOf { get; set; }
    }
}
