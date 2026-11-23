using System.Drawing;
using DevExpress.Docs.Spreadsheet;

class Program {
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Formulas.xlsx";
    static readonly FormulaType SelectedFormulaType = FormulaType.ReferencesInFormulas;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

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
        using var workbook = new Workbook();

        Worksheet dataWorksheet = workbook.Worksheets[0];
        dataWorksheet.Name = "Data";

        Worksheet formulasWorksheet = new Worksheet("Data And Formulas");
        workbook.Worksheets.Add(formulasWorksheet);

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

        workbook.Calculate();

        formulasWorksheet["E:G"].AutoFitColumns();

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream AddFormulaWithNamedRange(string rangeName, DocumentFormat outputFormat) {
        using var workbook = new Workbook();

        Worksheet formulasWorksheet = workbook.Worksheets[0];
        formulasWorksheet.Name = "Named Range In Formula";

        formulasWorksheet.Cells["B1"].Value = $"{rangeName}:";

        CellRange range = formulasWorksheet["A2:C5"];

        range.Value = 2;

        workbook.DefinedNames.Add(rangeName, range);

        range.ApplyBorders(new BorderFormat(new BorderLine(BorderLineStyle.Thin, Color.Black)));

        string formula = $"=SUM({rangeName})";
        formulasWorksheet.Cells["E1"].Value = "Formula";
        formulasWorksheet.Cells["F1"].Value = "Value";
        formulasWorksheet.Cells["E2"].Value = formula;

        formulasWorksheet.Cells["F2"].Formula = formula;

        workbook.Calculate();

        formulasWorksheet["E:F"].AutoFitColumns();

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream CreateNamedFormulas(string formulaName1, string formulaName2, DocumentFormat outputFormat) {
        using var workbook = new Workbook();

        Worksheet worksheet1 = workbook.Worksheets[0];
        worksheet1.Name = "Data";

        Worksheet worksheet2 = new Worksheet("Formulas");
        workbook.Worksheets.Add(worksheet2);
        worksheet2.Cells["A1"].Value = "Formula";
        worksheet2.Cells["B1"].Value = "Value";

        CellRange range = worksheet1["A1:C3"];

        range.Value = 2;

        worksheet1.DefinedNames.Add(formulaName1, "=SUM(Data!$A$1:$C$3)");

        workbook.DefinedNames.Add(formulaName2, $"=2*Data!{formulaName1}");

        worksheet2.Cells["A2"].Value = $"=Data!{formulaName1}";
        worksheet2.Cells["B2"].Formula = $"=Data!{formulaName1}";

        worksheet2.Cells["A3"].Value = $"={formulaName2}";
        worksheet2.Cells["B3"].Formula = $"={formulaName2}";

        worksheet2.Cells["A4"].Value = $"={formulaName2} + 100";
        worksheet2.Cells["B4"].Formula = $"={formulaName2} + 100";

        workbook.Calculate();

        worksheet2["A:B"].AutoFitColumns();

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream AddFormulasWithFunctions(DocumentFormat outputFormat) {
        using var workbook = new Workbook();

        Worksheet worksheet = workbook.Worksheets[0];
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
        worksheet.Cells["C6"].Format.NumberFormat = "m/d/yy";

        worksheet.Cells["A7"].Value = 15.12345;
        worksheet.Cells["B7"].Value = @"=UPPER(""formula"")";
        worksheet.Cells["C7"].Formula = @"=UPPER(""formula"")";

        workbook.Calculate();

        worksheet["A:C"].AutoFitColumns();

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream CreateSharedFormulas(DocumentFormat outputFormat) {
        using var workbook = new Workbook();

        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = "Shared Formulas";

        worksheet.Cells["A1"].Value = "Values";
        worksheet.Cells["A2"].Value = 1;

        worksheet["A3:A11"].Formula = "=SUM(A2+1)";

        workbook.Calculate();

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream CreateArrayFormulas(DocumentFormat outputFormat) {
        using var workbook = new Workbook();

        Worksheet worksheet1 = workbook.Worksheets[0];
        worksheet1.Name = "Legacy Array Formulas";

        worksheet1.Cells["A1"].Value = "Value 1";
        worksheet1.Cells["B1"].Value = "Value 2";
        worksheet1.Cells["D1"].Value = "Array Formula";
        worksheet1.Cells["E1"].Value = "Value";

        CellRange range = worksheet1["A2:B2"];
        range.Value = 2;

        range = worksheet1["A3:B3"];
        range.Value = 3;

        range = worksheet1["A4:B4"];
        range.Value = 4;

        worksheet1.Cells["D2"].Value = "=A2:A4*B2:B4";
        worksheet1["E2:E4"].ArrayFormula = "=A2:A4*B2:B4";

        Worksheet worksheet2 = new Worksheet("Dynamic Array Formulas");
        workbook.Worksheets.Add(worksheet2);
        worksheet2.Cells["A1"].DynamicArrayFormula = "={\"Red\",\"Green\",\"Orange\",\"Blue\"}";
        worksheet2.DynamicArrayFormulas.Add(new DynamicArrayFormula(worksheet2["A2"], "=LEN(A1:D1)"));

        workbook.Calculate();

        worksheet1["A:E"].AutoFitColumns();

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();

        workbook.Save(outputStream, outputFormat);

        outputStream.Position = 0;
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
