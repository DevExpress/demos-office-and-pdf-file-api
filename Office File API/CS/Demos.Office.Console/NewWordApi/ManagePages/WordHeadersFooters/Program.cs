using DevExpress.Docs.Word;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "HeaderFooter_template.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Modified_Document.docx";
    const string OutputFormat = "docx";
    const bool AddHeaders = true;
    const string HeaderText = "Document Title";
    const ParagraphAlignment HeaderAlignment = ParagraphAlignment.Center;
    const bool AddFooters = true;
    const string FooterText = "Confidential";
    const ParagraphAlignment FooterAlignment = ParagraphAlignment.Center;
    const bool DifferentFirstPage = true;
    const string FirstPageHeaderText = "Cover Page";
    const string FirstPageFooterText = "Internal Use Only";
    const ParagraphAlignment FirstPageHeaderAlignment = ParagraphAlignment.Center;
    const ParagraphAlignment FirstPageFooterAlignment = ParagraphAlignment.Center;
    const bool AddPageNumbering = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ProcessHeadersFootersDocument(inputStream, OutputFormat);
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

    static Stream ProcessHeadersFootersDocument(Stream templateStream, string outputFormat) {
        using var document = new WordDocument(templateStream);

        document.Sections[0].DifferentFirstPage = DifferentFirstPage;
        if(DifferentFirstPage) {
            AddHeader(document, HeaderFooterType.First, FirstPageHeaderText, FirstPageHeaderAlignment);
            AddFooter(document, HeaderFooterType.First, FirstPageFooterText, FirstPageFooterAlignment, false);
        }

        if(AddHeaders)
            AddHeader(document, HeaderFooterType.Odd, HeaderText, HeaderAlignment);
        if(AddFooters)
            AddFooter(document, HeaderFooterType.Odd, FooterText, FooterAlignment, AddPageNumbering);

        var outputStream = new MemoryStream();
        document.Save(outputStream, GetDocumentFormat(outputFormat));
        outputStream.Position = 0;
        return outputStream;
    }

    static void AddHeader(WordDocument document, HeaderFooterType headerFooterType, string text, ParagraphAlignment alignment) {
        foreach(var section in document.Sections) {
            var header = section.Headers[headerFooterType];
            if(header == null) {
                header = new Header();
                section.Headers[headerFooterType] = header;
            }

            if(!string.IsNullOrEmpty(text))
                header.InsertText(0, text);
            SetAlignment(header, alignment);
        }
    }

    static void AddFooter(WordDocument document, HeaderFooterType headerFooterType, string text, ParagraphAlignment alignment, bool addPageNumbering) {
        foreach(var section in document.Sections) {
            var footer = section.Footers[headerFooterType];
            if(footer == null) {
                footer = new Footer();
                section.Footers[headerFooterType] = footer;
            }

            if(!string.IsNullOrEmpty(text))
                footer.InsertText(0, text);
            if(addPageNumbering) {
                var paragraph = footer.AddParagraph("Page ");
                paragraph.InlineElements.Add(new Field("PAGE"));
                paragraph.AddText(" of ");
                paragraph.InlineElements.Add(new Field("NUMPAGES"));
            }

            SetAlignment(footer, alignment);
        }
    }

    static void SetAlignment(SectionHeaderFooterBase headerFooter, ParagraphAlignment alignment) {
        foreach(var paragraph in headerFooter.OfType<Paragraph>())
            paragraph.Properties.Alignment = alignment;
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
}