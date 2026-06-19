using DevExpress.Spreadsheet;
using System.Drawing;

class Program {
    const string OutputFileName = "Formulas.xlsx";
    static readonly FormulaType SelectedFormulaType = FormulaType.ReferencesInFormulas;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

            Stream outputStream = SelectedFormulaType switch {
                FormulaType.ReferencesInFormulas => AddFormulasWithCellReferences(OutputFormat),
                FormulaType.NamedRangesInFormulas => AddFormulaWithNamedRange("myRange", OutputFormat),
                FormulaType.CreateNamedFormulas => CreateNamedFormulas("Range_Sum", "Range_DoubleSum", OutputFormat),
                FormulaType.FunctionsInFormulas => AddFormulasWithFunctions(OutputFormat),
                FormulaType.CreateSharedFormulas => CreateSharedFormulas(OutputFormat),
                FormulaType.CreateArrayFormulas => CreateArrayFormulas(OutputFormat),
                _ => AddFormulasWithCellReferences(OutputFormat)
            };

            using(outputStream)
            using(var fileStream = File.Create(outputPath))
                outputStream.CopyTo(fileStream);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static Stream AddFormulasWithCellReferences(DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();

        Worksheet dataWorksheet = workbook.Worksheets.First();
        dataWorksheet.Name = "Data";

        Worksheet formulasWorksheet = workbook.Worksheets.Add();
        formulasWorksheet.Name = "Data And Formulas";

        dataWorksheet.Cells["A1"].Value = "Value 1";
        dataWorksheet.Cells["B1"].Value = "Value 2";
        dataWorksheet.Cells["C1"].Value = "Value 3";

        dataWorksheet.Cells["A2"].Value = 10;
        dataWorksheet.Cells["B2"].Value = 20;
        dataWorksheet.Cells["C2"].Value = 30;

        formulasWorksheet.Cells["A1"].Value = "Value 1";
        formulasWorksheet.Cells["B1"].Value = "Value 2";
        formulasWorksheet.Cells["C1"].Value = "Value 3";

        formulasWorksheet.Cells["A2"].Value = 10;
        formulasWorksheet.Cells["B2"].Value = 20;
        formulasWorksheet.Cells["C2"].Value = 30;

        formulasWorksheet.Cells["F1"].Value = "Formula";
        formulasWorksheet.Cells["G1"].Value = "Value";

        formulasWorksheet.Cells["E2"].Value = "A1 Relative References";
        formulasWorksheet.Cells["F2"].Value = "=A2+B2";
        formulasWorksheet.Cells["G2"].Formula = "=A2+B2";

        formulasWorksheet.Cells["E3"].Value = "A1 Absolute References";
        formulasWorksheet.Cells["F3"].Value = "=$A$2";
        formulasWorksheet.Cells["G3"].Formula = "=$A$2";

        formulasWorksheet.Cells["E4"].Value = "A1 Mixed References";
        formulasWorksheet.Cells["F4"].Value = "=A$2";
        formulasWorksheet.Cells["G4"].Formula = "=A$2";

        formulasWorksheet.Cells["E5"].Value = "Cross-Worksheet References";
        formulasWorksheet.Cells["F5"].Value = "=Data!A2 + Data!B2";
        formulasWorksheet.Cells["G5"].Formula = "=Data!A2 + Data!B2";

        formulasWorksheet.Cells["E6"].Value = "3D References";
        formulasWorksheet.Cells["F6"].Value = "=SUM('Data:Data And Formulas'!A2:C2)";
        formulasWorksheet.Cells["G6"].Formula = "=SUM('Data:Data And Formulas'!A2:C2)";

        formulasWorksheet.Columns.AutoFit("E", "G");
        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream AddFormulaWithNamedRange(string rangeName, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();

        Worksheet formulasWorksheet = workbook.Worksheets.First();
        formulasWorksheet.Name = "Named Range In Formula";

        formulasWorksheet["B1"].Value = $"{rangeName}:";

        CellRange range = formulasWorksheet.Range["A2:C5"];
        range.Value = 2;
        range.Name = "myRange";
        range.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin);

        string formula = $"=SUM({rangeName})";
        formulasWorksheet.Cells["E1"].Value = "Formula";
        formulasWorksheet.Cells["F1"].Value = "Value";
        formulasWorksheet.Cells["E2"].Value = formula;
        formulasWorksheet.Cells["F2"].Formula = formula;

        formulasWorksheet.Columns.AutoFit("E", "F");
        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream CreateNamedFormulas(string formulaName1, string formulaName2, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();

        Worksheet worksheet1 = workbook.Worksheets.First();
        worksheet1.Name = "Data";

        Worksheet worksheet2 = workbook.Worksheets.Add("Formulas");
        worksheet2.Cells["A1"].Value = "Formula";
        worksheet2.Cells["B1"].Value = "Value";

        CellRange range = worksheet1.Range["A1:C3"];
        range.Value = 2;

        worksheet1.DefinedNames.Add(formulaName1, "=SUM(Data!$A$1:$C$3)");
        workbook.DefinedNames.Add(formulaName2, $"=2*Data!{formulaName1}");

        worksheet2.Cells["A2"].Value = $"=Data!{formulaName1}";
        worksheet2.Cells["B2"].Formula = $"=Data!{formulaName1}";

        worksheet2.Cells["A3"].Value = $"={formulaName2}";
        worksheet2.Cells["B3"].Formula = $"={formulaName2}";

        worksheet2.Cells["A4"].Value = $"={formulaName2} + 100";
        worksheet2.Cells["B4"].Formula = $"={formulaName2} + 100";

        worksheet2.Columns.AutoFit("A", "B");
        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream AddFormulasWithFunctions(DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();

        Worksheet worksheet = workbook.Worksheets.First();
        worksheet.Name = "Functions In Formulas";

        worksheet.Cells["A1"].Value = "Data";
        worksheet.Cells["B1"].Value = "Formula";
        worksheet.Cells["C1"].Value = "Value";

        worksheet.Cells["A2"].Value = 15;
        worksheet.Cells["B2"].Value = @"=IF(A2<10, ""Normal"", ""Excess"")";
        worksheet.Cells["C2"].Formula = @"=IF(A2<10, ""Normal"", ""Excess"")";

        worksheet.Cells["A3"].Value = 3;
        worksheet.Cells["B3"].Value = "=AVERAGE(A2:A7)";
        worksheet.Cells["C3"].Formula = "=AVERAGE(A2:A7)";

        worksheet.Cells["A4"].Value = 3;
        worksheet.Cells["B4"].Value = "=SUM(A3:A5,A6,100)";
        worksheet.Cells["C4"].Formula = "=SUM(A3:A5,A6,100)";

        worksheet.Cells["A5"].Value = 3;
        worksheet.Cells["B5"].Value = "=ROUND(SUM(A6,A7),2)";
        worksheet.Cells["C5"].Formula = "=ROUND(SUM(A6,A7),2)";

        worksheet.Cells["A6"].Value = 20;
        worksheet.Cells["B6"].Value = "=Today()";
        worksheet.Cells["C6"].Formula = "=Today()";
        worksheet.Cells["C6"].NumberFormat = "m/d/yy";

        worksheet.Cells["A7"].Value = 15.12345;
        worksheet.Cells["B7"].Value = @"=UPPER(""formula"")";
        worksheet.Cells["C7"].Formula = @"=UPPER(""formula"")";

        worksheet.Columns.AutoFit("A", "C");
        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream CreateSharedFormulas(DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();

        Worksheet worksheet = workbook.Worksheets.First();
        worksheet.Name = "Shared Formulas";

        worksheet.Cells["A1"].Value = "Values";
        worksheet.Cells["A2"].Value = 1;
        worksheet.Range["A3:A11"].Formula = "=SUM(A2+1)";

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream CreateArrayFormulas(DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();

        Worksheet worksheet1 = workbook.Worksheets.First();
        worksheet1.Name = "Legacy Array Formulas";

        worksheet1.Cells["A1"].Value = "Value 1";
        worksheet1.Cells["B1"].Value = "Value 2";
        worksheet1.Cells["D1"].Value = "Array Formula";
        worksheet1.Cells["E1"].Value = "Value";

        CellRange range = worksheet1.Range["A2:B2"];
        range.Value = 2;

        range = worksheet1.Range["A3:B3"];
        range.Value = 3;

        range = worksheet1.Range["A4:B4"];
        range.Value = 4;

        worksheet1.Range["D2"].Value = "=A2:A4*B2:B4";
        worksheet1.Range["E2:E4"].ArrayFormula = "=A2:A4*B2:B4";

        worksheet1.Columns.AutoFit("A", "E");

        Worksheet worksheet2 = workbook.Worksheets.Add("Dynamic Array Formulas");
        worksheet2["A1"].DynamicArrayFormulaInvariant = "={\"Red\",\"Green\",\"Orange\",\"Blue\"}";
        worksheet2.DynamicArrayFormulas.Add(worksheet2["A2"], "=LEN(A1:D1)");

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }

    enum FormulaType {
        ReferencesInFormulas,
        NamedRangesInFormulas,
        CreateNamedFormulas,
        FunctionsInFormulas,
        CreateSharedFormulas,
        CreateArrayFormulas
    }
}
