using DevExpress.Office;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Fields_Template.docx");
    static readonly string LogoPath = Path.Combine("Documents", "DevExpress.png");
    const string OutputFileName = "Fields_Result.docx";
    const string OutputFormat = "docx";
    const FieldAction Action = FieldAction.Insert;

    static readonly string[] FieldEntryValues = {
        "DATE",
        "TIME",
        "PAGE",
        "NUMPAGES",
        "DocVariable",
        "AUTHOR",
        "TITLE",
        "DOCPROPERTY"
    };

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

            using var templateStream = File.OpenRead(inputPath);
            using var outputStream = Action switch {
                FieldAction.Insert => InsertFields(templateStream, FieldEntryValues, GetDocumentFormat(OutputFormat)),
                FieldAction.Update => UpdateFields(templateStream, GetDocumentFormat(OutputFormat)),
                FieldAction.Lock => LockFields(templateStream, FieldEntryValues, GetDocumentFormat(OutputFormat)),
                FieldAction.Unlink => UnlinkFields(templateStream, FieldEntryValues, GetDocumentFormat(OutputFormat)),
                FieldAction.Delete => DeleteFields(templateStream, FieldEntryValues, GetDocumentFormat(OutputFormat)),
                _ => throw new InvalidOperationException($"Unsupported action: {Action}")
            };
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

    static MemoryStream InsertFields(Stream stream, IEnumerable<string> fieldNames, DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(stream);
        Document document = wordProcessor.Document;
        Table table = document.Tables.First();
        document.CalculateDocumentVariable += CalculateDocumentVariable;
        foreach(var fieldType in fieldNames) {
            var row = table.Rows.Append();
            var firstCell = row.FirstCell;
            var secondCell = row.LastCell;
            switch(fieldType) {
                case "DATE":
                    document.InsertText(firstCell.Range.Start, "New Date Field");
                    document.Fields.Create(secondCell.Range.Start, "DATE");
                    break;
                case "TIME":
                    document.InsertText(firstCell.Range.Start, "New Time Field");
                    document.Fields.Create(secondCell.Range.Start, "TIME");
                    break;
                case "PAGE":
                    document.InsertText(firstCell.Range.Start, "New Page Field");
                    document.Fields.Create(secondCell.Range.Start, "PAGE");
                    break;
                case "NUMPAGES":
                    document.InsertText(firstCell.Range.Start, "New NumPages Field");
                    document.Fields.Create(secondCell.Range.Start, "NUMPAGES");
                    break;
                case "DocVariable":
                    document.InsertText(firstCell.Range.Start, "New DocVariable Field");
                    document.Fields.Create(secondCell.Range.Start, "DocVariable MyVariable");
                    break;
                case "AUTHOR":
                    document.InsertText(firstCell.Range.Start, "New Author Field");
                    document.DocumentProperties.Author = "Test Author";
                    document.Fields.Create(secondCell.Range.Start, "AUTHOR");
                    break;
                case "TITLE":
                    document.InsertText(firstCell.Range.Start, "New Title Field");
                    document.DocumentProperties.Title = "Test Title";
                    document.Fields.Create(secondCell.Range.Start, "TITLE");
                    break;
                case "DOCPROPERTY":
                    document.InsertText(firstCell.Range.Start, "New DocProperty Field");
                    document.CustomProperties.Add("MyProperty", "Custom Property Value");
                    document.Fields.Create(secondCell.Range.Start, "DOCPROPERTY MyProperty");
                    break;
            }
        }
        document.UpdateAllFields();
        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static MemoryStream UpdateFields(Stream stream, DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.CalculateDocumentVariable += CalculateDocumentVariable;
        wordProcessor.LoadDocument(stream);
        Document document = wordProcessor.Document;
        document.UpdateAllFields();
        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static MemoryStream LockFields(Stream stream, IEnumerable<string> fieldNames, DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(stream);
        Document document = wordProcessor.Document;
        document.CalculateDocumentVariable += CalculateDocumentVariable;
        document.UpdateAllFields();
        foreach(Field field in document.Fields) {
            field.ShowCodes = true;
            string fieldCode = document.GetText(field.CodeRange).ToUpper().TrimStart();
            if(fieldNames.Any(fn => fieldCode.StartsWith(fn, StringComparison.InvariantCultureIgnoreCase)))
                field.Locked = true;
        }
        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static MemoryStream UnlinkFields(Stream stream, IEnumerable<string> fieldNames, DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(stream);
        Document document = wordProcessor.Document;
        document.CalculateDocumentVariable += CalculateDocumentVariable;
        document.UpdateAllFields();
        for(int i = document.Fields.Count - 1; i >= 0; i--) {
            Field field = document.Fields[i];
            field.ShowCodes = true;
            string fieldCode = document.GetText(field.CodeRange).ToUpper().TrimStart();
            if(fieldNames.Any(fn => fieldCode.StartsWith(fn, StringComparison.InvariantCultureIgnoreCase)))
                field.Unlink();
        }
        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static MemoryStream DeleteFields(Stream stream, IEnumerable<string> fieldNames, DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(stream);
        Document document = wordProcessor.Document;
        document.CalculateDocumentVariable += CalculateDocumentVariable;
        document.UpdateAllFields();
        for(int i = document.Fields.Count - 1; i >= 0; i--) {
            Field field = document.Fields[i];
            field.ShowCodes = true;
            string fieldCode = document.GetText(field.CodeRange).ToUpper().TrimStart();
            if(fieldNames.Any(fn => fieldCode.StartsWith(fn, StringComparison.InvariantCultureIgnoreCase)))
                document.Delete(field.Range);
        }
        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static void CalculateDocumentVariable(object? sender, CalculateDocumentVariableEventArgs e) {
        if(e.VariableName == "MyVariable") {
            e.Value = "My Test Variable " + DateTime.Now.ToString("MMMM dd, yyyy");
            e.Handled = true;
        }
        if(e.VariableName == "Test") {
            var tempWordProcessor = new RichEditDocumentServer();
            tempWordProcessor.Unit = DocumentUnit.Inch;
            Document tempDoc = tempWordProcessor.Document;
            Shape picture = tempDoc.Shapes.InsertPicture(tempDoc.Range.Start, DocumentImageSource.FromFile(Path.Combine(AppContext.BaseDirectory, LogoPath)));
            picture.TextWrapping = TextWrappingType.InLineWithText;
            picture.LockAspectRatio = true;
            picture.Width = 2f;
            e.Value = tempWordProcessor;
            e.Handled = true;
        }
    }

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "doc" => DocumentFormat.Doc,
            "docx" => DocumentFormat.Docx,
            "docm" => DocumentFormat.Docm,
            "dotx" => DocumentFormat.Dotx,
            "dotm" => DocumentFormat.Dotm,
            "rtf" => DocumentFormat.Rtf,
            _ => DocumentFormat.Docx
        };
    }

    enum FieldAction {
        Insert,
        Update,
        Lock,
        Unlink,
        Delete
    }
}
