# DevExpress Office & PDF File API: .NET Console App Demos

[DevExpress Office & PDF File API](https://www.devexpress.com/products/net/office-file-api/) is a non-visual .NET library that allows you to generate, read, and modify PDF files and Microsoft Office documents (Word, Excel, PowerPoint)[^1].

[^1]: Office, Excel, Word, PowerPoint are trademarks or registered trademarks of Microsoft Corporation.

## What's Inside

This folder includes source code for DevExpress Office & PDF File API demos (.NET Console Apps). If you used DevExpress Unified Installer, you can find the same code in your [local installation folder](https://docs.devexpress.com/OfficeFileAPI/15097/demo-application#find-the-source-code).

## Build and Run Demo Apps

Each example in this repository is an individual console project. See the list below for project descriptions and corresponding project paths.

### Windows

Locate a project you need to run. You can use the following command:

```powershell
Get-ChildItem -Recurse -Filter *.csproj | Select-Object -ExpandProperty FullName
```

Open a terminal window in the `Demos.Office/Console` folder. Run a module with the `--project` parameter:

```bash
dotnet run --project ./PdfProcessing/DocumentConversion/PdfToImage/PdfToImage.csproj
```

Review output files in the module folder (`bin/Debug/...`) or copy the path displayed by the demo.

### Linux & macOS

Locate a project you need to run. You can use the following command:

```bash
find . -name "*.csproj"
```

Most modules are cross-platform and only require .NET SDK. Additional dependencies are required in the following cases:

- `System.Drawing.Common`/GDI+-based scenarios:
	- Linux: install `libgdiplus` and a basic font package.
	- macOS: install `mono-libgdiplus` (Homebrew) and ensure that fonts are available.
- Digital signature demos require valid certificates/keys on your machine.

Typical install commands:

```bash
# Ubuntu/Debian
sudo apt-get update
sudo apt-get install -y libgdiplus fontconfig

# macOS (Homebrew)
brew install mono-libgdiplus
```

Open a terminal window in the `Demos.Office/Console` folder. Run a module with the `--project` parameter:

```bash
dotnet run --project ./PdfProcessing/DocumentConversion/PdfToImage/PdfToImage.csproj
```

If a demo fails on Linux/macOS, verify native dependencies and fonts and run:

```bash
dotnet --info
dotnet restore
dotnet build
```

## Included Demos: AI-powered Extensions

These demos use [AI-powered Extensions for DevExpress Office & PDF File API](https://docs.devexpress.com/OfficeFileAPI/405645/ai-powered-extensions) for intelligent document processing tasks.

Before running AI demos, review prerequisites: [AI-powered Extensions prerequisites](https://docs.devexpress.com/OfficeFileAPI/405645/ai-powered-extensions#prerequisites).

| Demo | Folder | Program.cs |
|---|---|---|
| Summarize a PDF Document | [PdfSummarize](./AI/DocumentOperations/PdfSummarize) | [Program.cs](./AI/DocumentOperations/PdfSummarize/Program.cs) |
| Summarize a Presentation | [PresentationSummarize](./AI/DocumentOperations/PresentationSummarize) | [Program.cs](./AI/DocumentOperations/PresentationSummarize/Program.cs) |
| Translate a Presentation | [PresentationTranslate](./AI/DocumentOperations/PresentationTranslate) | [Program.cs](./AI/DocumentOperations/PresentationTranslate/Program.cs) |
| Proofread a Word Document | [WordProofread](./AI/DocumentOperations/WordProofread) | [Program.cs](./AI/DocumentOperations/WordProofread/Program.cs) |
| Summarize a Word Document | [WordSummarize](./AI/DocumentOperations/WordSummarize) | [Program.cs](./AI/DocumentOperations/WordSummarize/Program.cs) |
| Translate a Word Document | [WordTranslate](./AI/DocumentOperations/WordTranslate) | [Program.cs](./AI/DocumentOperations/WordTranslate/Program.cs) |

## Included Demos: New PDF Documents API (CTP)

Our [new PDF Document API](https://docs.devexpress.com/OfficeFileAPI/405907/pdf-document-api-new/overview) implements a strongly typed document object model (DOM) and allows you to access and modify all parts of a document, including pages, content fragments, annotations, form fields, metadata, attachments, and logical structure elements.

| Demo | Folder | Program.cs |
|---|---|---|
| Add PDF Annotations | [PdfAnnotations](./NewPdfDocumentApi/DocumentElements/PdfAnnotations) | [Program.cs](./NewPdfDocumentApi/DocumentElements/PdfAnnotations/Program.cs) |
| Attach Files to a PDF Document | [PdfAttachments](./NewPdfDocumentApi/DocumentElements/PdfAttachments) | [Program.cs](./NewPdfDocumentApi/DocumentElements/PdfAttachments/Program.cs) |
| Add a Watermark to a PDF Document | [PdfApplyWatermark](./NewPdfDocumentApi/DocumentProtection/PdfApplyWatermark) | [Program.cs](./NewPdfDocumentApi/DocumentProtection/PdfApplyWatermark/Program.cs) |
| Encrypt a PDF Document | [PdfEncryptPassword](./NewPdfDocumentApi/DocumentProtection/PdfEncryptPassword) | [Program.cs](./NewPdfDocumentApi/DocumentProtection/PdfEncryptPassword/Program.cs) |
| Redact a PDF Document | [PdfRedaction](./NewPdfDocumentApi/DocumentProtection/PdfRedaction) | [Program.cs](./NewPdfDocumentApi/DocumentProtection/PdfRedaction/Program.cs) |
| Create Interactive PDF Forms | [PdfCreateForm](./NewPdfDocumentApi/Forms/PdfCreateForm) | [Program.cs](./NewPdfDocumentApi/Forms/PdfCreateForm/Program.cs) |
| Generate a PDF Invoice | [PdfGenerateInvoice](./NewPdfDocumentApi/GenerateDocument/PdfGenerateInvoice) | [Program.cs](./NewPdfDocumentApi/GenerateDocument/PdfGenerateInvoice/Program.cs) |
| Remove Pages from a PDF Document | [PdfClearPages](./NewPdfDocumentApi/ManagePages/PdfClearPages) | [Program.cs](./NewPdfDocumentApi/ManagePages/PdfClearPages/Program.cs) |
| Reorder, Insert, and Delete PDF Pages | [PdfManagePages](./NewPdfDocumentApi/ManagePages/PdfManagePages) | [Program.cs](./NewPdfDocumentApi/ManagePages/PdfManagePages/Program.cs) |
| Merge PDF Documents | [PdfMergeDocuments](./NewPdfDocumentApi/MergeSplit/PdfMergeDocuments) | [Program.cs](./NewPdfDocumentApi/MergeSplit/PdfMergeDocuments/Program.cs) |
| Split a PDF Document | [PdfSplitDocuments](./NewPdfDocumentApi/MergeSplit/PdfSplitDocuments) | [Program.cs](./NewPdfDocumentApi/MergeSplit/PdfSplitDocuments/Program.cs) |
| Extract Images from a PDF Document | [PdfExtractImages](./NewPdfDocumentApi/SearchExtract/PdfExtractImages) | [Program.cs](./NewPdfDocumentApi/SearchExtract/PdfExtractImages/Program.cs) |
| Find and Highlight Text in a PDF Document | [PdfFindHighlight](./NewPdfDocumentApi/SearchExtract/PdfFindHighlight) | [Program.cs](./NewPdfDocumentApi/SearchExtract/PdfFindHighlight/Program.cs) |
| Extract PDF XMP Metadata | [PdfXMPMetadata](./NewPdfDocumentApi/SearchExtract/PdfXMPMetadata) | [Program.cs](./NewPdfDocumentApi/SearchExtract/PdfXMPMetadata/Program.cs) |

## Included Demos: PDF Documents API

[DevExpress PDF Document API](https://www.devexpress.com/products/net/office-file-api/pdf/) allows you to generate, edit, convert, merge, split, password-protect, and digitally sign PDF files.

For additional information, refer to [DevExpress PDF Document API Documentation](https://docs.devexpress.com/OfficeFileAPI/16491/pdf-document-api).

| Demo | Folder | Program.cs |
|---|---|---|
| Convert a PDF Document to PDF/A | [PdfToCompatible](./PdfProcessing/DocumentConversion/PdfToCompatible) | [Program.cs](./PdfProcessing/DocumentConversion/PdfToCompatible/Program.cs) |
| Convert a PDF Document to Images | [PdfToImage](./PdfProcessing/DocumentConversion/PdfToImage) | [Program.cs](./PdfProcessing/DocumentConversion/PdfToImage/Program.cs) |
| Add a Watermark to a PDF Document | [PdfApplyWatermark](./PdfProcessing/DocumentProtection/PdfApplyWatermark) | [Program.cs](./PdfProcessing/DocumentProtection/PdfApplyWatermark/Program.cs) |
| Sign a PDF Document | [PdfDigitalSignature](./PdfProcessing/DocumentProtection/PdfDigitalSignature) | [Program.cs](./PdfProcessing/DocumentProtection/PdfDigitalSignature/Program.cs) |
| Encrypt a PDF Document | [PdfEncryptPassword](./PdfProcessing/DocumentProtection/PdfEncryptPassword) | [Program.cs](./PdfProcessing/DocumentProtection/PdfEncryptPassword/Program.cs) |
| Redact a PDF Document | [PdfRedaction](./PdfProcessing/DocumentProtection/PdfRedaction) | [Program.cs](./PdfProcessing/DocumentProtection/PdfRedaction/Program.cs) |
| Validate Signatures in a PDF Document | [PdfValidateSignature](./PdfProcessing/DocumentProtection/PdfValidateSignature) | [Program.cs](./PdfProcessing/DocumentProtection/PdfValidateSignature/Program.cs) |
| Split a PDF Document | [PdfSplitDocuments](./PdfProcessing/MergeSplit/PdfSplitDocuments) | [Program.cs](./PdfProcessing/MergeSplit/PdfSplitDocuments/Program.cs) |
| Add Annotations to a PDF Document | [PdfAnnotations](./PdfProcessing/DocumentElements/PdfAnnotations) | [Program.cs](./PdfProcessing/DocumentElements/PdfAnnotations/Program.cs) |
| Manage Attachments in a PDF Document | [PdfAttachments](./PdfProcessing/DocumentElements/PdfAttachments) | [Program.cs](./PdfProcessing/DocumentElements/PdfAttachments/Program.cs) |
| Add ZUGFeRD Invoice to a PDF Document | [PdfAttachZUGFeRDInvoice](./PdfProcessing/DocumentElements/PdfAttachZUGFeRDInvoice) | [Program.cs](./PdfProcessing/DocumentElements/PdfAttachZUGFeRDInvoice/Program.cs) |
| Manage Bookmarks in a PDF Document | [PdfBookmarks](./PdfProcessing/DocumentElements/PdfBookmarks) | [Program.cs](./PdfProcessing/DocumentElements/PdfBookmarks/Program.cs) |
| Create a PDF AcroForm | [PdfCreateForm](./PdfProcessing/Forms/PdfCreateForm) | [Program.cs](./PdfProcessing/Forms/PdfCreateForm/Program.cs) |
| Fill a PDF Form | [PdfFillForm](./PdfProcessing/Forms/PdfFillForm) | [Program.cs](./PdfProcessing/Forms/PdfFillForm/Program.cs) |
| Import and Export Form Data | [PdfImportExportFormData](./PdfProcessing/Forms/PdfImportExportFormData) | [Program.cs](./PdfProcessing/Forms/PdfImportExportFormData/Program.cs) |
| Remove and Flatten a PDF Form | [PdfRemoveFlattenForm](./PdfProcessing/Forms/PdfRemoveFlattenForm) | [Program.cs](./PdfProcessing/Forms/PdfRemoveFlattenForm/Program.cs) |
| Clear Pages in a PDF Document | [PdfClearPages](./PdfProcessing/ManagePages/PdfClearPages) | [Program.cs](./PdfProcessing/ManagePages/PdfClearPages/Program.cs) |
| Draw on a PDF Page | [PdfDrawPage](./PdfProcessing/ManagePages/PdfDrawPage) | [Program.cs](./PdfProcessing/ManagePages/PdfDrawPage/Program.cs) |
| Manage Pages in a PDF Document | [PdfManagePages](./PdfProcessing/ManagePages/PdfManagePages) | [Program.cs](./PdfProcessing/ManagePages/PdfManagePages/Program.cs) |
| Rotate and Resize PDF Pages | [PdfRotateResizePages](./PdfProcessing/ManagePages/PdfRotateResizePages) | [Program.cs](./PdfProcessing/ManagePages/PdfRotateResizePages/Program.cs) |
| Merge PDF Documents | [PdfMergeDocuments](./PdfProcessing/MergeSplit/PdfMergeDocuments) | [Program.cs](./PdfProcessing/MergeSplit/PdfMergeDocuments/Program.cs) |
| Extract Images from a PDF Document | [PdfExtractImages](./PdfProcessing/SearchExtract/PdfExtractImages) | [Program.cs](./PdfProcessing/SearchExtract/PdfExtractImages/Program.cs) |
| Extract Text from a PDF Document | [PdfExtractText](./PdfProcessing/SearchExtract/PdfExtractText) | [Program.cs](./PdfProcessing/SearchExtract/PdfExtractText/Program.cs) |
| Find and Highlight Text in a PDF Document | [PdfFindHighlight](./PdfProcessing/SearchExtract/PdfFindHighlight) | [Program.cs](./PdfProcessing/SearchExtract/PdfFindHighlight/Program.cs) |
| Extract PDF Document Metadata | [PdfXMPMetadata](./PdfProcessing/SearchExtract/PdfXMPMetadata) | [Program.cs](./PdfProcessing/SearchExtract/PdfXMPMetadata/Program.cs) |

## Included Demos: Word Document API

[DevExpress Word (RTF) Document API](https://www.devexpress.com/products/net/office-file-api/word/) allows you to create, read, edit, convert, merge, split, password-protect, and digitally sign Microsoft Word documents.

For additional information, refer to [Word Processing Document API Documentation](https://docs.devexpress.com/OfficeFileAPI/17488/word-processing-document-api).

| Demo | Folder | Program.cs |
|---|---|---|
| Convert Word Document Pages to Images | [WordToImage](./WordProcessing/DocumentConversion/WordToImage) | [Program.cs](./WordProcessing/DocumentConversion/WordToImage/Program.cs) |
| Convert a Word Document to PDF | [WordToPdf](./WordProcessing/DocumentConversion/WordToPdf) | [Program.cs](./WordProcessing/DocumentConversion/WordToPdf/Program.cs) |
| Convert a Word Document to Another Format | [WordToWord](./WordProcessing/DocumentConversion/WordToWord) | [Program.cs](./WordProcessing/DocumentConversion/WordToWord/Program.cs) |
| Add a Digital Signature to a Word Document | [WordDigitalSignature](./WordProcessing/DocumentProtection/WordDigitalSignature) | [Program.cs](./WordProcessing/DocumentProtection/WordDigitalSignature/Program.cs) |
| Encrypt a Word Document | [WordEncryptPassword](./WordProcessing/DocumentProtection/WordEncryptPassword) | [Program.cs](./WordProcessing/DocumentProtection/WordEncryptPassword/Program.cs) |
| Restrict Word Document Modifications | [WordProtectPassword](./WordProcessing/DocumentProtection/WordProtectPassword) | [Program.cs](./WordProcessing/DocumentProtection/WordProtectPassword/Program.cs) |
| Mail Merge | [WordMailMerge](./WordProcessing/MailMergeFields/WordMailMerge) | [Program.cs](./WordProcessing/MailMergeFields/WordMailMerge/Program.cs) |
| Define Master-Detail Regions in a Template | [WordMailMergeRegions](./WordProcessing/MailMergeFields/WordMailMergeRegions) | [Program.cs](./WordProcessing/MailMergeFields/WordMailMergeRegions/Program.cs) |
| Merge Word Documents | [Merge](./WordProcessing/MergeSplit/Merge) | [Program.cs](./WordProcessing/MergeSplit/Merge/Program.cs) |
| Split a Word Document | [Split](./WordProcessing/MergeSplit/Split) | [Program.cs](./WordProcessing/MergeSplit/Split/Program.cs) |
| Work with Barcodes in a Word Document | [WordBarcodes](./WordProcessing/DocumentElements/WordBarcodes) | [Program.cs](./WordProcessing/DocumentElements/WordBarcodes/Program.cs) |
| Work with Charts in a Word Document | [WordCharts](./WordProcessing/DocumentElements/WordCharts) | [Program.cs](./WordProcessing/DocumentElements/WordCharts/Program.cs) |
| Work with Content Controls in a Word Document | [WordContentControls](./WordProcessing/DocumentElements/WordContentControls) | [Program.cs](./WordProcessing/DocumentElements/WordContentControls/Program.cs) |
| Work with Footnotes and Endnotes in a Word Document | [WordFootnotesEndnotes](./WordProcessing/DocumentElements/WordFootnotesEndnotes) | [Program.cs](./WordProcessing/DocumentElements/WordFootnotesEndnotes/Program.cs) |
| Work with Lists in a Word Document | [WordLists](./WordProcessing/DocumentElements/WordLists) | [Program.cs](./WordProcessing/DocumentElements/WordLists/Program.cs) |
| Work with OLE Objects in a Word Document | [WordOleObjects](./WordProcessing/DocumentElements/WordOleObjects) | [Program.cs](./WordProcessing/DocumentElements/WordOleObjects/Program.cs) |
| Work with Shapes in a Word Document | [WordShapes](./WordProcessing/DocumentElements/WordShapes) | [Program.cs](./WordProcessing/DocumentElements/WordShapes/Program.cs) |
| Work with Tables in a Word Document | [WordTables](./WordProcessing/DocumentElements/WordTables) | [Program.cs](./WordProcessing/DocumentElements/WordTables/Program.cs) |
| Inspect and Sanitize a Word Document | [WordInspectAndSanitize](./WordProcessing/DocumentProtection/WordInspectAndSanitize) | [Program.cs](./WordProcessing/DocumentProtection/WordInspectAndSanitize/Program.cs) |
| Secure Word Document Loading | [WordSecureDocumentLoading](./WordProcessing/DocumentProtection/WordSecureDocumentLoading) | [Program.cs](./WordProcessing/DocumentProtection/WordSecureDocumentLoading/Program.cs) |
| Format Paragraphs in a Word Document | [WordParagraphFormatting](./WordProcessing/Formatting/WordParagraphFormatting) | [Program.cs](./WordProcessing/Formatting/WordParagraphFormatting/Program.cs) |
| Change Table Styles in a Word Document | [WordTableStyle](./WordProcessing/Formatting/WordTableStyle) | [Program.cs](./WordProcessing/Formatting/WordTableStyle/Program.cs) |
| Format Text in a Word Document | [WordTextFormatting](./WordProcessing/Formatting/WordTextFormatting) | [Program.cs](./WordProcessing/Formatting/WordTextFormatting/Program.cs) |
| Work with Fields in a Word Document | [WordFields](./WordProcessing/MailMergeFields/WordFields) | [Program.cs](./WordProcessing/MailMergeFields/WordFields/Program.cs) |
| Work with Table of Contents in a Word Document | [WordTableOfContents](./WordProcessing/MailMergeFields/WordTableOfContents) | [Program.cs](./WordProcessing/MailMergeFields/WordTableOfContents/Program.cs) |
| Work with Format Sections in a Word Document | [WordFormatSections](./WordProcessing/ManagePages/WordFormatSections) | [Program.cs](./WordProcessing/ManagePages/WordFormatSections/Program.cs) |
| Work with Headers and Footers in a Word Document | [WordHeadersFooters](./WordProcessing/ManagePages/WordHeadersFooters) | [Program.cs](./WordProcessing/ManagePages/WordHeadersFooters/Program.cs) |
| Work with Watermarks in a Word Document | [WordWatermark](./WordProcessing/ManagePages/WordWatermark) | [Program.cs](./WordProcessing/ManagePages/WordWatermark/Program.cs) |
| Work with Comments in a Word Document | [WordComments](./WordProcessing/ReviewComment/WordComments) | [Program.cs](./WordProcessing/ReviewComment/WordComments/Program.cs) |
| Compare Word Documents | [WordCompareDocuments](./WordProcessing/ReviewComment/WordCompareDocuments) | [Program.cs](./WordProcessing/ReviewComment/WordCompareDocuments/Program.cs) |
| Track Changes in a Word Document | [WordTrackChanges](./WordProcessing/ReviewComment/WordTrackChanges) | [Program.cs](./WordProcessing/ReviewComment/WordTrackChanges/Program.cs) |
| Obtain Word Document Statistics | [WordDocumentStatistics](./WordProcessing/SearchExtract/WordDocumentStatistics) | [Program.cs](./WordProcessing/SearchExtract/WordDocumentStatistics/Program.cs) |
| Extract Images from a Word Document | [WordExtractImages](./WordProcessing/SearchExtract/WordExtractImages) | [Program.cs](./WordProcessing/SearchExtract/WordExtractImages/Program.cs) |
| Find and Replace Text in a Word Document | [WordFindReplace](./WordProcessing/SearchExtract/WordFindReplace) | [Program.cs](./WordProcessing/SearchExtract/WordFindReplace/Program.cs) |
| Extract Word Document Metadata | [WordMetadata](./WordProcessing/SearchExtract/WordMetadata) | [Program.cs](./WordProcessing/SearchExtract/WordMetadata/Program.cs) |

## Included Demos: Excel Spreadsheet API

[DevExpress Excel Spreadsheet API for .NET](https://www.devexpress.com/products/net/office-file-api/spreadsheet/) allows you to create, read, write, convert, print, and protect Microsoft Excel files.

For additional information, refer to [Spreadsheet Document API Documentation](https://docs.devexpress.com/OfficeFileAPI/14912/spreadsheet-document-api).

| Demo | Folder | Program.cs |
|---|---|---|
| Calculate Formulas in a Workbook | [ExcelCalculateWorkbook](./SpreadsheetProcessing/Calculations/ExcelCalculateWorkbook) | [Program.cs](./SpreadsheetProcessing/Calculations/ExcelCalculateWorkbook/Program.cs) |
| Detect Circular References in a Workbook | [ExcelCircularReferences](./SpreadsheetProcessing/Calculations/ExcelCircularReferences) | [Program.cs](./SpreadsheetProcessing/Calculations/ExcelCircularReferences/Program.cs) |
| Calculate Custom Functions in a Workbook | [ExcelCustomFunctions](./SpreadsheetProcessing/Calculations/ExcelCustomFunctions) | [Program.cs](./SpreadsheetProcessing/Calculations/ExcelCustomFunctions/Program.cs) |
| Calculate Excel Formulas in a Workbook | [ExcelFormulas](./SpreadsheetProcessing/Calculations/ExcelFormulas) | [Program.cs](./SpreadsheetProcessing/Calculations/ExcelFormulas/Program.cs) |
| Export a Data Table to Excel | [ExcelExportDataTable](./SpreadsheetProcessing/DataImportExport/ExcelExportDataTable) | [Program.cs](./SpreadsheetProcessing/DataImportExport/ExcelExportDataTable/Program.cs) |
| Import a Data Table from Excel | [ExcelImportDataTable](./SpreadsheetProcessing/DataImportExport/ExcelImportDataTable) | [Program.cs](./SpreadsheetProcessing/DataImportExport/ExcelImportDataTable/Program.cs) |
| Import Data from a List | [ExcelImportList](./SpreadsheetProcessing/DataImportExport/ExcelImportList) | [Program.cs](./SpreadsheetProcessing/DataImportExport/ExcelImportList/Program.cs) |
| Convert an Excel Spreadsheet to Another Excel Format | [ExcelToExcel](./SpreadsheetProcessing/DocumentConversion/ExcelToExcel) | [Program.cs](./SpreadsheetProcessing/DocumentConversion/ExcelToExcel/Program.cs) |
| Convert an Excel Spreadsheet to HTML | [ExcelToHtml](./SpreadsheetProcessing/DocumentConversion/ExcelToHtml) | [Program.cs](./SpreadsheetProcessing/DocumentConversion/ExcelToHtml/Program.cs) |
| Convert an Excel Spreadsheet to Images | [ExcelToImage](./SpreadsheetProcessing/DocumentConversion/ExcelToImage) | [Program.cs](./SpreadsheetProcessing/DocumentConversion/ExcelToImage/Program.cs) |
| Convert an Excel Spreadsheet to PDF | [ExcelToPdf](./SpreadsheetProcessing/DocumentConversion/ExcelToPdf) | [Program.cs](./SpreadsheetProcessing/DocumentConversion/ExcelToPdf/Program.cs) |
| Work with Charts in Excel Spreadsheets | [ExcelCharts](./SpreadsheetProcessing/DocumentElements/ExcelCharts) | [Program.cs](./SpreadsheetProcessing/DocumentElements/ExcelCharts/Program.cs) |
| Work with Excel OLE Objects | [ExcelOleObjects](./SpreadsheetProcessing/DocumentElements/ExcelOleObjects) | [Program.cs](./SpreadsheetProcessing/DocumentElements/ExcelOleObjects/Program.cs) |
| Work with Excel Pivot Tables | [ExcelPivotTables](./SpreadsheetProcessing/DocumentElements/ExcelPivotTables) | [Program.cs](./SpreadsheetProcessing/DocumentElements/ExcelPivotTables/Program.cs) |
| Work with Excel Shapes | [ExcelShapes](./SpreadsheetProcessing/DocumentElements/ExcelShapes) | [Program.cs](./SpreadsheetProcessing/DocumentElements/ExcelShapes/Program.cs) |
| Work with Excel Tables | [ExcelTables](./SpreadsheetProcessing/DocumentElements/ExcelTables) | [Program.cs](./SpreadsheetProcessing/DocumentElements/ExcelTables/Program.cs) |
| Work with Excel Threaded Comments | [ExcelThreadedComments](./SpreadsheetProcessing/DocumentElements/ExcelThreadedComments) | [Program.cs](./SpreadsheetProcessing/DocumentElements/ExcelThreadedComments/Program.cs) |
| Digitally Sign an Excel Document | [ExcelDigitalSignature](./SpreadsheetProcessing/DocumentProtection/ExcelDigitalSignature) | [Program.cs](./SpreadsheetProcessing/DocumentProtection/ExcelDigitalSignature/Program.cs) |
| Encrypt an Excel Document with a Password | [ExcelEncryptPassword](./SpreadsheetProcessing/DocumentProtection/ExcelEncryptPassword) | [Program.cs](./SpreadsheetProcessing/DocumentProtection/ExcelEncryptPassword/Program.cs) |
| Inspect and Sanitize an Excel Document | [ExcelInspectAndSanitize](./SpreadsheetProcessing/DocumentProtection/ExcelInspectAndSanitize) | [Program.cs](./SpreadsheetProcessing/DocumentProtection/ExcelInspectAndSanitize/Program.cs) |
| Protect an Excel Workbook | [ExcelProtectWorkbook](./SpreadsheetProcessing/DocumentProtection/ExcelProtectWorkbook) | [Program.cs](./SpreadsheetProcessing/DocumentProtection/ExcelProtectWorkbook/Program.cs) |
| Protect an Excel Worksheet | [ExcelProtectWorksheet](./SpreadsheetProcessing/DocumentProtection/ExcelProtectWorksheet) | [Program.cs](./SpreadsheetProcessing/DocumentProtection/ExcelProtectWorksheet/Program.cs) |
| Load an Excel Document Securely | [ExcelSecureDocumentLoading](./SpreadsheetProcessing/DocumentProtection/ExcelSecureDocumentLoading) | [Program.cs](./SpreadsheetProcessing/DocumentProtection/ExcelSecureDocumentLoading/Program.cs) |
| Format Cells in an Excel Document | [ExcelCellFormatting](./SpreadsheetProcessing/Formatting/ExcelCellFormatting) | [Program.cs](./SpreadsheetProcessing/Formatting/ExcelCellFormatting/Program.cs) |
| Apply Conditional Formatting in an Excel Document | [ExcelConditionalFormatting](./SpreadsheetProcessing/Formatting/ExcelConditionalFormatting) | [Program.cs](./SpreadsheetProcessing/Formatting/ExcelConditionalFormatting/Program.cs) |
| Apply Rich Text Formatting in an Excel Document | [ExcelRichTextFormatting](./SpreadsheetProcessing/Formatting/ExcelRichTextFormatting) | [Program.cs](./SpreadsheetProcessing/Formatting/ExcelRichTextFormatting/Program.cs) |
| Clear Contents in an Excel Worksheet | [ExcelClearContents](./SpreadsheetProcessing/ManageSheets/ExcelClearContents) | [Program.cs](./SpreadsheetProcessing/ManageSheets/ExcelClearContents/Program.cs) |
| Validate Data in an Excel Worksheet | [ExcelDataValidation](./SpreadsheetProcessing/ManageSheets/ExcelDataValidation) | [Program.cs](./SpreadsheetProcessing/ManageSheets/ExcelDataValidation/Program.cs) |
| Filter Data in an Excel Worksheet | [ExcelFilterData](./SpreadsheetProcessing/ManageSheets/ExcelFilterData) | [Program.cs](./SpreadsheetProcessing/ManageSheets/ExcelFilterData/Program.cs) |
| Format Rows and Columns in an Excel Worksheet | [ExcelFormatRowsColumns](./SpreadsheetProcessing/ManageSheets/ExcelFormatRowsColumns) | [Program.cs](./SpreadsheetProcessing/ManageSheets/ExcelFormatRowsColumns/Program.cs) |
| Freeze Panes in an Excel Worksheet | [ExcelFreezePanes](./SpreadsheetProcessing/ManageSheets/ExcelFreezePanes) | [Program.cs](./SpreadsheetProcessing/ManageSheets/ExcelFreezePanes/Program.cs) |
| Group Data in an Excel Worksheet | [ExcelGroupData](./SpreadsheetProcessing/ManageSheets/ExcelGroupData) | [Program.cs](./SpreadsheetProcessing/ManageSheets/ExcelGroupData/Program.cs) |
| Insert Cell Values in an Excel Worksheet | [ExcelInsertCellValues](./SpreadsheetProcessing/ManageSheets/ExcelInsertCellValues) | [Program.cs](./SpreadsheetProcessing/ManageSheets/ExcelInsertCellValues/Program.cs) |
| Set Print Settings in an Excel Worksheet | [ExcelPrintSettings](./SpreadsheetProcessing/ManageSheets/ExcelPrintSettings) | [Program.cs](./SpreadsheetProcessing/ManageSheets/ExcelPrintSettings/Program.cs) |
| Sort Data in an Excel Worksheet | [ExcelSortData](./SpreadsheetProcessing/ManageSheets/ExcelSortData) | [Program.cs](./SpreadsheetProcessing/ManageSheets/ExcelSortData/Program.cs) |
| Add Worksheets in an Excel Workbook | [ExcelWorksheetOperations](./SpreadsheetProcessing/ManageSheets/ExcelWorksheetOperations) | [Program.cs](./SpreadsheetProcessing/ManageSheets/ExcelWorksheetOperations/Program.cs) |
| Merge Excel Workbooks | [ExcelMergeWorkbooks](./SpreadsheetProcessing/MergeSplitDocuments/ExcelMergeWorkbooks) | [Program.cs](./SpreadsheetProcessing/MergeSplitDocuments/ExcelMergeWorkbooks/Program.cs) |
| Split an Excel Workbook | [ExcelSplitWorkbooks](./SpreadsheetProcessing/MergeSplitDocuments/ExcelSplitWorkbooks) | [Program.cs](./SpreadsheetProcessing/MergeSplitDocuments/ExcelSplitWorkbooks/Program.cs) |
| Find and Replace Text in an Excel Document | [ExcelFindReplace](./SpreadsheetProcessing/SearchExtract/ExcelFindReplace) | [Program.cs](./SpreadsheetProcessing/SearchExtract/ExcelFindReplace/Program.cs) |
| Extract Text from an Excel Document | [ExcelMetadata](./SpreadsheetProcessing/SearchExtract/ExcelMetadata) | [Program.cs](./SpreadsheetProcessing/SearchExtract/ExcelMetadata/Program.cs) |
| Obtain Excel Document Statistics | [ExcelWorkbookStatistics](./SpreadsheetProcessing/SearchExtract/ExcelWorkbookStatistics) | [Program.cs](./SpreadsheetProcessing/SearchExtract/ExcelWorkbookStatistics/Program.cs) |

## Included Demos: PowerPoint Presentation API

[DevExpress PowerPoint Presentation API for .NET](https://www.devexpress.com/products/net/office-file-api/powerpoint-presentation-api/) allows you to generate, edit, merge, print/export, and digitally sign Microsoft PowerPoint files.

For additional information, refer to [DevExpress Presentation API Documentation](https://docs.devexpress.com/OfficeFileAPI/405405/presentation-api/overview).

| Demo | Folder | Program.cs |
|---|---|---|
| Convert a Presentation to Images | [PresentationToImage](./PowerPointPresentation/DocumentConversion/PresentationToImage) | [Program.cs](./PowerPointPresentation/DocumentConversion/PresentationToImage/Program.cs) |
| Add Charts to a Presentation | [PresentationCharts](./PowerPointPresentation/DocumentElements/PresentationCharts) | [Program.cs](./PowerPointPresentation/DocumentElements/PresentationCharts/Program.cs) |
| Add Shapes to a Presentation | [PresentationShapes](./PowerPointPresentation/DocumentElements/PresentationShapes) | [Program.cs](./PowerPointPresentation/DocumentElements/PresentationShapes/Program.cs) |
| Convert a Presentation to PDF | [PresentationToPdf](./PowerPointPresentation/DocumentConversion/PresentationToPdf) | [Program.cs](./PowerPointPresentation/DocumentConversion/PresentationToPdf/Program.cs) |
| Convert a Presentation to Another Format | [PresentationToPresentation](./PowerPointPresentation/DocumentConversion/PresentationToPresentation) | [Program.cs](./PowerPointPresentation/DocumentConversion/PresentationToPresentation/Program.cs) |
| Add Speaker Notes to a Presentation | [PresentationNotes](./PowerPointPresentation/DocumentElements/PresentationNotes) | [Program.cs](./PowerPointPresentation/DocumentElements/PresentationNotes/Program.cs) |
| Add a Table to a Presentation Slide | [PresentationTables](./PowerPointPresentation/DocumentElements/PresentationTables) | [Program.cs](./PowerPointPresentation/DocumentElements/PresentationTables/Program.cs) |
| Merge Presentations | [PresentationMerge](./PowerPointPresentation/MergeSplit/PresentationMerge) | [Program.cs](./PowerPointPresentation/MergeSplit/PresentationMerge/Program.cs) |
| Split a Presentation | [PresentationSplit](./PowerPointPresentation/MergeSplit/PresentationSplit) | [Program.cs](./PowerPointPresentation/MergeSplit/PresentationSplit/Program.cs) |
| Extract Images from a Presentation | [PresentationExtractImages](./PowerPointPresentation/SearchExtract/PresentationExtractImages) | [Program.cs](./PowerPointPresentation/SearchExtract/PresentationExtractImages/Program.cs) |
| Extract Text from a Presentation | [PresentationExtractText](./PowerPointPresentation/SearchExtract/PresentationExtractText) | [Program.cs](./PowerPointPresentation/SearchExtract/PresentationExtractText/Program.cs) |
| Find and Replace Text in a Presentation | [PresentationFindReplace](./PowerPointPresentation/SearchExtract/PresentationFindReplace) | [Program.cs](./PowerPointPresentation/SearchExtract/PresentationFindReplace/Program.cs) |
| Extract Presentation Metadata | [PresentationMetadata](./PowerPointPresentation/SearchExtract/PresentationMetadata) | [Program.cs](./PowerPointPresentation/SearchExtract/PresentationMetadata/Program.cs) |

## Online Demos

- [Office File API Online Demos](https://demos.devexpress.com/OfficeFileAPI/BlazorServer/)

## More Examples

- [Word Processing Document API - How to Process Word Documents in Code](https://github.com/DevExpress-Examples/word-document-api-examples)
- [Spreadsheet Document API - How to Process Spreadsheet Documents in Code (Part 1)](https://github.com/DevExpress-Examples/spreadsheet-document-api-examples-part1)
- [Spreadsheet Document API - How to Process Spreadsheet Documents in Code (Part 2)](https://github.com/DevExpress-Examples/spreadsheet-document-api-examples-part2)