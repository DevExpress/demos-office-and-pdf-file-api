using DevExpress.Docs.Word;
using DevExpress.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Fields_Template.docx");
    static readonly string LogoPath = Path.Combine("Documents", "DevExpress.png");
    const string OutputDirectoryName = "Output";
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
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var templateStream = File.OpenRead(inputPath);
            using var outputStream = Action switch {
                FieldAction.Insert => InsertFields(templateStream, FieldEntryValues, OutputFormat),
                FieldAction.Update => UpdateFields(templateStream, OutputFormat),
                FieldAction.Lock => LockFields(templateStream, FieldEntryValues, OutputFormat),
                FieldAction.Unlink => UnlinkFields(templateStream, FieldEntryValues, OutputFormat),
                FieldAction.Delete => DeleteFields(templateStream, FieldEntryValues, OutputFormat),
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

    static MemoryStream InsertFields(Stream stream, IEnumerable<string> fieldNames, string outputFormat) {
        using var document = new WordDocument(stream);
        var table = document.Body.OfType<Table>().First();
        foreach(var fieldType in fieldNames) {
            switch(fieldType) {
                case "DATE":
                    AppendFieldRow(table, "New Date Field", "DATE");
                    break;
                case "TIME":
                    AppendFieldRow(table, "New Time Field", "TIME");
                    break;
                case "PAGE":
                    AppendFieldRow(table, "New Page Field", "PAGE");
                    break;
                case "NUMPAGES":
                    AppendFieldRow(table, "New NumPages Field", "NUMPAGES");
                    break;
                case "DocVariable":
                    AppendFieldRow(table, "New DocVariable Field", "DocVariable MyVariable");
                    break;
                case "AUTHOR":
                    document.DocumentProperties.Author = "Test Author";
                    AppendFieldRow(table, "New Author Field", "AUTHOR");
                    break;
                case "TITLE":
                    document.DocumentProperties.Title = "Test Title";
                    AppendFieldRow(table, "New Title Field", "TITLE");
                    break;
                case "DOCPROPERTY":
                    document.DocumentProperties.CustomProperties["MyProperty"] = "Custom Property Value";
                    AppendFieldRow(table, "New DocProperty Field", "DOCPROPERTY MyProperty");
                    break;
            }
        }

        UpdateAllFields(document);
        return SaveDocument(document, outputFormat);
    }

    static void AppendFieldRow(Table table, string caption, string fieldCode) {
        var row = table.Rows.Append();
        row.Cells.First.Content.Add(new Paragraph(caption));
        var fieldParagraph = new Paragraph();
        fieldParagraph.InlineElements.Add(new Field(fieldCode));
        row.Cells.Last.Content.Add(fieldParagraph);
    }

    static MemoryStream UpdateFields(Stream stream, string outputFormat) {
        using var document = new WordDocument(stream);
        UpdateAllFields(document);
        return SaveDocument(document, outputFormat);
    }

    static MemoryStream LockFields(Stream stream, IEnumerable<string> fieldNames, string outputFormat) {
        using var document = new WordDocument(stream);
        UpdateAllFields(document);
        foreach(var paragraph in GetParagraphs(document.Body)) {
            foreach(var field in paragraph.InlineElements.OfType<Field>()) {
                if(HasFieldType(field, fieldNames))
                    field.Locked = true;
            }
        }

        return SaveDocument(document, outputFormat);
    }

    static MemoryStream UnlinkFields(Stream stream, IEnumerable<string> fieldNames, string outputFormat) {
        using var document = new WordDocument(stream);
        UpdateAllFields(document);
        foreach(var paragraph in GetParagraphs(document.Body)) {
            var inlineElements = paragraph.InlineElements;
            for(int i = inlineElements.Count - 1; i >= 0; i--) {
                if(inlineElements[i] is Field field && HasFieldType(field, fieldNames)) {
                    var resultElements = field.Result.OfType<Inline>().Select(element => (Inline)element.Clone()).ToList();
                    inlineElements.RemoveAt(i);
                    for(int j = 0; j < resultElements.Count; j++)
                        inlineElements.Insert(i + j, resultElements[j]);
                }
            }
        }

        return SaveDocument(document, outputFormat);
    }

    static MemoryStream DeleteFields(Stream stream, IEnumerable<string> fieldNames, string outputFormat) {
        using var document = new WordDocument(stream);
        UpdateAllFields(document);
        foreach(var paragraph in GetParagraphs(document.Body)) {
            var inlineElements = paragraph.InlineElements;
            for(int i = inlineElements.Count - 1; i >= 0; i--) {
                if(inlineElements[i] is Field field && HasFieldType(field, fieldNames))
                    inlineElements.RemoveAt(i);
            }
        }

        return SaveDocument(document, outputFormat);
    }

    static List<Paragraph> GetParagraphs(IEnumerable<Block> blocks) {
        var result = new List<Paragraph>();
        foreach(var block in blocks) {
            if(block is Paragraph paragraph)
                result.Add(paragraph);
            else if(block is Table table) {
                foreach(var row in table.Rows) {
                    foreach(var cell in row.Cells)
                        result.AddRange(GetParagraphs(cell.Content));
                }
            }
        }

        return result;
    }

    static bool HasFieldType(Field field, IEnumerable<string> fieldNames) {
        var fieldCode = field.Code.GetText().TrimStart();
        return fieldNames.Any(fieldName => fieldCode.StartsWith(fieldName, StringComparison.InvariantCultureIgnoreCase));
    }

    static void UpdateAllFields(WordDocument document) {
        document.UpdateAllFields(new FieldUpdateOptions { DocumentVariableProvider = new DemoDocumentVariableProvider() });
    }

    static MemoryStream SaveDocument(WordDocument document, string outputFormat) {
        var outputStream = new MemoryStream();
        document.Save(outputStream, GetDocumentFormat(outputFormat));
        outputStream.Position = 0;
        return outputStream;
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

    class DemoDocumentVariableProvider : IDocumentVariableProvider {
        const float LogoWidth = 600f;

        public DocumentVariableValue? GetValue(string variableName, IList<string> arguments) {
            if(variableName == "MyVariable")
                return new DocumentVariableValue("My Test Variable " + DateTime.Now.ToString("MMMM dd, yyyy"));
            if(variableName == "Test")
                return new DocumentVariableValue(CreateLogoDocument());
            return null;
        }

        WordDocument CreateLogoDocument() {
            var image = DXImage.FromStream(new MemoryStream(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, LogoPath))));
            var picture = new PictureShape(image) {
                LockAspectRatio = true,
                Width = LogoWidth,
                Height = LogoWidth * image.Height / image.Width
            };
            var result = new WordDocument();
            result.Body.Add(picture);
            return result;
        }
    }

    enum FieldAction {
        Insert,
        Update,
        Lock,
        Unlink,
        Delete
    }
}