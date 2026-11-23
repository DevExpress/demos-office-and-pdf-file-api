using DevExpress.Docs.Word;
using DevExpress.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "MailMerge_Template.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "MailMerge_Result.pdf";
    const string OutputFormat = "pdf";
    const bool IsPdfOutput = true;
    const int EmployeeId = 1;
    const string CustomerId = "ALFKI";
    const string EmployeePhotoPrefix = "dbimg://";

    static readonly Employee[] Employees = {
        new(1, "Nancy", "Davolio", "507 - 20th Ave. E.\nApt. 2A", "Seattle", "98122", "Sales Representative", "(206) 555-9857")
    };

    static readonly Customer[] Customers = {
        new("ALFKI", "Alfreds Futterkiste", "Maria Anders", "Sales Representative", "Obere Str. 57", "Berlin", "12209")
    };

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            var selectedEmployee = Employees.FirstOrDefault(employee => employee.EmployeeID == EmployeeId);
            if(selectedEmployee != null) {
                var selectedCustomer = Customers.FirstOrDefault(customer => customer.CustomerID == CustomerId);
                if(selectedCustomer != null)
                    selectedEmployee.Customer = selectedCustomer;
            }

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = MailMergeDocument(inputStream, selectedEmployee, IsPdfOutput, OutputFormat)
                ?? throw new InvalidOperationException("Employee data was not found.");
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

    static MemoryStream? MailMergeDocument(Stream stream, Employee? selectedEmployee, bool isPdfOutput, string outputFormat) {
        using var document = new WordDocument(stream);

        if(selectedEmployee == null)
            return null;

        ReplaceEmployeePhotoFields(document.Body, selectedEmployee.EmployeeID);

        var options = new MailMergeOptions {
            DataSource = new List<Employee>() { selectedEmployee },
            FirstRecordIndex = 0,
            LastRecordIndex = 0
        };

        using var mergedDocument = document.MailMerge(options);

        var outputStream = new MemoryStream();
        if(isPdfOutput)
            mergedDocument.ExportToPdf(outputStream);
        else
            mergedDocument.Save(outputStream, GetDocumentFormat(outputFormat));

        outputStream.Position = 0;
        return outputStream;
    }

    static void ReplaceEmployeePhotoFields(IEnumerable<Block> blocks, int employeeId) {
        foreach(var block in blocks) {
            if(block is Paragraph paragraph)
                ReplaceEmployeePhotoFields(paragraph, employeeId);
            else if(block is Table table) {
                foreach(var row in table.Rows) {
                    foreach(var cell in row.Cells)
                        ReplaceEmployeePhotoFields(cell.Content, employeeId);
                }
            }
        }
    }

    static void ReplaceEmployeePhotoFields(Paragraph paragraph, int employeeId) {
        var inlineElements = paragraph.InlineElements;
        for(int i = 0; i < inlineElements.Count; i++) {
            if(inlineElements[i] is Field field && IsEmployeePhotoField(field))
                inlineElements[i] = CreateEmployeePhoto(employeeId);
        }
    }

    static bool IsEmployeePhotoField(Field field) {
        var fieldCode = field.Code.GetText().TrimStart();
        return fieldCode.StartsWith("INCLUDEPICTURE", StringComparison.OrdinalIgnoreCase)
               && fieldCode.Contains(EmployeePhotoPrefix, StringComparison.OrdinalIgnoreCase);
    }

    static PictureShape CreateEmployeePhoto(int employeeId) {
        var fileName = Path.Combine(AppContext.BaseDirectory, "Documents", $"Photo{employeeId}.jpeg");
        var imageStream = new MemoryStream(File.ReadAllBytes(fileName));
        var result = new PictureShape(DXImage.FromStream(imageStream));
        return result;
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

    class Employee {
        public Employee(int employeeId, string firstName, string lastName, string address, string city, string postalCode, string title, string homePhone, byte[]? picture = null) {
            EmployeeID = employeeId;
            FirstName = firstName;
            LastName = lastName;
            Address = address;
            City = city;
            PostalCode = postalCode;
            Title = title;
            HomePhone = homePhone;
            Picture = picture;
        }

        public int EmployeeID { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string FullName => $"{FirstName} {LastName}";
        public string Title { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string PostalCode { get; set; }
        public string HomePhone { get; set; }
        public Customer? Customer { get; set; }
        public byte[]? Picture { get; set; }
    }

    class Customer {
        public Customer(string customerId, string companyName, string contactName, string contactTitle, string address, string city, string postalCode) {
            CustomerID = customerId;
            CompanyName = companyName;
            ContactName = contactName;
            ContactTitle = contactTitle;
            Address = address;
            City = city;
            PostalCode = postalCode;
        }

        public string CustomerID { get; set; }
        public string CompanyName { get; set; }
        public string ContactName { get; set; }
        public string ContactTitle { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string PostalCode { get; set; }
    }
}