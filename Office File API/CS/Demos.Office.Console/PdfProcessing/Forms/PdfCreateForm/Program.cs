using DevExpress.Drawing;
using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_BlankLetter.pdf");
    const string OutputDirectoryName = "Output";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var outputStream = CreateForm(File.OpenRead(inputPath), createTextBox: true, createComboBox: true, createListBox: true, createCheckBox: true, createRadioGroup: true);
            SaveDocument(outputStream, Path.Combine(outputDirectory, "Document.pdf"));

            Console.WriteLine($"Created {outputDirectory}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static void SaveDocument(Stream stream, string outputPath) {
        stream.Position = 0;
        using var fileStream = File.Create(outputPath);
        stream.CopyTo(fileStream);
    }

    static Stream CreateForm(Stream documentStream, bool createTextBox, bool createComboBox, bool createListBox, bool createCheckBox, bool createRadioGroup) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);
        var formPage = processor.Document.Pages[0];

        using var g = processor.CreateGraphicsPageSystem();
        using var brush = new DXSolidBrush(System.Drawing.Color.Black);
        var titleFont = new DXFont("Arial", 25);
        var labelFont = new DXFont("Arial", 12);

        var stringSize = g.MeasureString("Sample Form", titleFont);
        g.DrawString("Sample Form", titleFont, brush, new System.Drawing.PointF((float)formPage.MediaBox.Width / 2 - stringSize.Width / 2, (float)formPage.MediaBox.Top - 50));

        List<PdfAcroFormField> formFields = new();

        if(createTextBox) {
            g.DrawString("Text Box :", labelFont, brush, new System.Drawing.PointF((float)formPage.MediaBox.Left + 10, 520));

            var textBoxField = PdfAcroFormField.CreateTextBox("TextBox", 1, new PdfRectangle(formPage.MediaBox.Right - 510, 500, formPage.MediaBox.Right - 10, 520));
            textBoxField.Text = "Sample Text";
            textBoxField.Appearance.FontSize = 12;

            formFields.Add(textBoxField);
        }

        if(createComboBox) {
            g.DrawString("Combo Box :", labelFont, brush, new System.Drawing.PointF((float)formPage.MediaBox.Left + 10, 480));

            var comboBoxField = PdfAcroFormField.CreateComboBox("ComboBox", 1, new PdfRectangle(formPage.MediaBox.Right - 510, 460, formPage.MediaBox.Right - 10, 480));
            comboBoxField.AddValue("Item 1");
            comboBoxField.AddValue("Item 2");
            comboBoxField.AddValue("Item 3");
            comboBoxField.SelectValue("Item 1");
            comboBoxField.Appearance.FontSize = 12;

            formFields.Add(comboBoxField);
        }

        if(createListBox) {
            g.DrawString("List Box :", labelFont, brush, new System.Drawing.PointF((float)formPage.MediaBox.Left + 10, 440));

            var listBoxField = PdfAcroFormField.CreateListBox("ListBox", 1, new PdfRectangle(formPage.MediaBox.Right - 510, 380, formPage.MediaBox.Right - 10, 440));
            listBoxField.AddValue("Item 1");
            listBoxField.AddValue("Item 2");
            listBoxField.AddValue("Item 3");
            listBoxField.SelectValue("Item 1");
            listBoxField.Appearance.FontSize = 12;

            formFields.Add(listBoxField);
        }

        if(createCheckBox) {
            g.DrawString("Check Box :", labelFont, brush, new System.Drawing.PointF((float)formPage.MediaBox.Left + 10, 360));

            var checkBoxField = PdfAcroFormField.CreateCheckBox("CheckBox", 1, new PdfRectangle(formPage.MediaBox.Right - 510, 340, formPage.MediaBox.Right - 490, 360));
            checkBoxField.IsChecked = true;

            formFields.Add(checkBoxField);
        }

        if(createRadioGroup) {
            g.DrawString("Radio Group :", labelFont, brush, new System.Drawing.PointF((float)formPage.MediaBox.Left + 10, 320));

            var fieldRadioGroup = PdfAcroFormField.CreateRadioGroup("RadioGroup", 1);
            fieldRadioGroup.AddButton("RadioOption1", new PdfRectangle(formPage.MediaBox.Right - 510, 300, formPage.MediaBox.Right - 490, 320));
            fieldRadioGroup.AddButton("RadioOption2", new PdfRectangle(formPage.MediaBox.Right - 470, 300, formPage.MediaBox.Right - 450, 320));
            fieldRadioGroup.SelectedIndex = 0;

            formFields.Add(fieldRadioGroup);
        }

        g.AddToPageForeground(formPage);
        processor.AddFormFields(formFields);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }
}
