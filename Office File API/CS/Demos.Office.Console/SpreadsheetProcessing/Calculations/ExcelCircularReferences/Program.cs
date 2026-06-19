using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;
using DevExpress.XtraSpreadsheet.Services;
using System.Text;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "CircularReferencesSample.xlsx");
    const string OutputFileName = "Circular_References.txt";
    const int MaxIterationCount = 5;

    static async Task<int> Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            var calculationResult = await DetectWorkbookCircularReferences(inputStream, MaxIterationCount, DocumentFormat.Text);
            await File.WriteAllTextAsync(outputPath, calculationResult, Encoding.UTF8);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static async Task<string> DetectWorkbookCircularReferences(Stream inputStream, int maxIterationCount, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        await workbook.LoadDocumentAsync(inputStream);

        CustomCalculationService customCalculationService = new CustomCalculationService();
        workbook.ReplaceService<ICustomCalculationService>(customCalculationService);

        if(workbook.Options.CalculationEngineType != CalculationEngineType.ChainBased)
            workbook.Options.CalculationEngineType = CalculationEngineType.ChainBased;

        workbook.DocumentSettings.Calculation.Iterative = true;
        workbook.DocumentSettings.Calculation.MaxIterationCount = maxIterationCount;
        workbook.CalculateFullRebuild();

        var circularReferencedCells = customCalculationService.CircularReferencedCells;
        CalculationResult calculationResult = new CalculationResult();
        foreach(var circularReferencedCell in circularReferencedCells) {
            var worksheet = workbook.Worksheets[circularReferencedCell.SheetId - 1];
            string sheetName = worksheet.Name;

            var cell = worksheet.Cells[circularReferencedCell.RowIndex, circularReferencedCell.ColumnIndex];
            string cellAddress = cell.GetReferenceA1();
            object cellValue = cell.Value.ToObject();
            calculationResult.Cells.Add(new CellInfo(sheetName, cellAddress, cellValue));
        }

        return calculationResult.ToString();
    }

    public class CustomCalculationService : ICustomCalculationService {
        public HashSet<CellKey> CircularReferencedCells { get; set; } = new HashSet<CellKey>();
        public CustomCalculationService() { }
        void Reset() {
            CircularReferencedCells.Clear();
        }

        public bool OnBeginCalculation() {
            Reset();
            return true;
        }

        public void OnBeginCellCalculation(CellCalculationArgs args) { }

        public bool OnBeginCircularReferencesCalculation() {
            return true;
        }

        public void OnEndCalculation() { }

        public void OnEndCellCalculation(CellKey cellKey, CellValue startValue, CellValue endValue) { }

        public void OnEndCircularReferencesCalculation(IList<CellKey> cellKeys) {
            CircularReferencedCells = new HashSet<CellKey>(cellKeys);
        }

        public bool ShouldMarkupCalculateAlwaysCells() {
            return true;
        }
    }

    public class CalculationResult {
        public bool HasCircularReferences { get { return Cells.Count > 0; } }
        public List<CellInfo> Cells { get; set; } = new List<CellInfo>();
        public override string ToString() {
            StringBuilder sb = new StringBuilder();
            if(HasCircularReferences) {
                sb.AppendLine("Circular references detected.");
                sb.AppendLine(string.Join(Environment.NewLine, Cells));
            } else {
                sb.Append("No circular references detected.");
            }
            return sb.ToString().TrimEnd();
        }
    }

    public class CellInfo {
        public string SheetName { get; set; }
        public string A1Reference { get; set; }
        public object Value { get; set; }
        public CellInfo(string sheetName, string a1Reference, object value) {
            SheetName = sheetName;
            A1Reference = a1Reference;
            Value = value;
        }
        public override string ToString() {
            return $"Sheet = \"{SheetName}\", Cell = {A1Reference}, Value = {Value}";
        }
    }
}
