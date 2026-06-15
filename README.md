# DevExpress Office & PDF File API: Process DOCX Documents, XLSX Spreadsheets, PPTX Presentations, and PDF Files in Blazor Web Apps

The DevExpress Office & PDF File API library includes demo source code. You can find it in the current folder or in your [local installation folder](https://docs.devexpress.com/OfficeFileAPI/15097/demo-application#find-the-source-code).

## Build Blazor Demos for Office & PDF File API

To run demos on your machine, you must:

1. Clone the current repository:

    HTTPS:
    ```
    git clone https://github.com/DevExpress/demos-office-and-pdf-file-api.git
    ```
    SSH:
    ```
    git clone git@github.com:DevExpress/demos-office-and-pdf-file-api.git
    ```

2. Open the following solution file in Visual Studio:

    * _Office File API/CS/Demos.Office.Web/Demos.Office.Web.sln_

3. The following [NuGet packages](https://docs.devexpress.com/OfficeFileAPI/400744/installation-guide/install-nuget-packages) are already referenced in project files:

	* [Markdig](https://nuget.org/packages/Markdig)
	* [System.Drawing.Common](https://nuget.org/packages/System.Drawing.Common)
	* [DevExpress.Blazor](https://nuget.org/packages/DevExpress.Blazor)
	* [DevExpress.AIIntegration.Docs](https://nuget.org/packages/DevExpress.AIIntegration.Docs)
	* [DevExpress.Docs.Pdf](https://nuget.org/packages/DevExpress.Docs.Pdf)
	* [DevExpress.Docs.Presentation](https://nuget.org/packages/DevExpress.Docs.Presentation)
	* [DevExpress.Document.Processor](https://nuget.org/packages/DevExpress.Document.Processor)

    You do not need to install them manually.

4. Run the application.

## AI Setup (Requirements and Registration)

Review general requirements for AI-powered demos: [AI-powered Extensions for DevExpress Office File API](https://docs.devexpress.com/OfficeFileAPI/405645/ai-powered-extensions).

- [AI-powered Extensions for DevExpress Office File API](https://docs.devexpress.com/OfficeFileAPI/405645/ai-powered-extensions)

In this project, AI configuration and service registration are defined in [Program.cs](./Program.cs):

- Read AI settings from configuration (`AIIntegrationSettings`): [Program.cs#L32-L34](./Program.cs#L32-L34)
- Register chat client in DI: [Program.cs#L47](./Program.cs#L47)
- Register DevExpress AI document processing service: [Program.cs#L50-L51](./Program.cs#L50-L51)

Expected configuration keys are loaded from the `AIIntegrationSettings` section:

- `EndpointUrl`
- `Key`
- `DeploymentName`

## Project Structure (Demos.Office.Web)

The Blazor demo application is organized as follows:

- [App.razor](./App.razor), [Program.cs](./Program.cs), [Routes.razor](./Routes.razor): app startup, DI configuration, and routing.
- [Pages](./Pages): Blazor demo pages (UI and page-level logic) grouped by tasks:
  - [Pages/AI](./Pages/AI)
  - [Pages/NewPdfDocumentApi](./Pages/NewPdfDocumentApi)
  - [Pages/PdfProcessing](./Pages/PdfProcessing)
  - [Pages/WordProcessing](./Pages/WordProcessing)
  - [Pages/SpreadsheetProcessing](./Pages/SpreadsheetProcessing)
  - [Pages/Presentation](./Pages/Presentation)
- [Controllers](./Controllers): HTTP endpoints and backend processing logic for demos grouped by features:
  - [Controllers/AI](./Controllers/AI)
  - [Controllers/NewPdfDocumentApi](./Controllers/NewPdfDocumentApi)
  - [Controllers/PdfProcessing](./Controllers/PdfProcessing)
  - [Controllers/WordProcessing](./Controllers/WordProcessing)
  - [Controllers/SpreadsheetProcessing](./Controllers/SpreadsheetProcessing)
  - [Controllers/PowerPointPresentation](./Controllers/PowerPointPresentation)
- [Models](./Models): demo-specific DTOs/view models and feature metadata.
- [Services](./Services): shared helper/services used by processing demos (format conversion, extraction, utility logic).
- [DemoShell](./DemoShell): reusable layout/navigation/search UI components for the demo host.
- [Documents](./Documents): source files/templates used in demo scenarios.
- [wwwroot](./wwwroot): static assets.

## Locate Office & PDF File API Code

If you want to locate code directly related to DevExpress Office & PDF File API types and methods, start from the following locations:

1. Controller actions in [Controllers](./Controllers):

    This is the primary backend execution layer, where document processing operations are implemented (load/process/save/export). Check feature subfolders first (for example, [Controllers/PdfProcessing](./Controllers/PdfProcessing), [Controllers/WordProcessing](./Controllers/WordProcessing), [Controllers/SpreadsheetProcessing](./Controllers/SpreadsheetProcessing)).

2. Demo pages in [Pages](./Pages):

    These pages define user workflows, collect parameters, and invoke controller endpoints. Review corresponding feature folders to trace UI-to-backend flow.

3. Supporting models/services:

    - [Models](./Models): request/response shapes, file format descriptors, and demo metadata.
    - [Services](./Services): reusable helpers used across multiple demos.

4. Sample input files in [Documents](./Documents):

    Many demos use files from this folder as input templates or test data.

### Recommended Navigation Flow

1. Open the corresponding page in [Pages](./Pages).
2. In [Controllers](./Controllers), find the endpoint/action that this page calls.
3. Review related helper code in [Services](./Services) and data contracts in [Models](./Models).
4. Review input/output artifacts in [Documents](./Documents).


## Online Demos

* [DevExpress Office & PDF File API Demos for Blazor](https://demos.devexpress.com/OfficeFileApi/)

## Helpful Resources

* [Documentation](https://docs.devexpress.com/OfficeFileAPI/14911/office-file-api)
* [Examples](https://github.com/search?q=topic%3Aoffice-file-api+org%3ADevExpress-Examples&type=repositories)
* [Videos](https://www.youtube.com/playlist?list=PL8h4jt35t1wg72KSHpBJ9VDElNCYdlWMC)
* [Support Center](https://supportcenter.devexpress.com/ask)
