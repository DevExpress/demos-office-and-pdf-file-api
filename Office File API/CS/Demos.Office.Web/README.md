# DevExpress Office & PDF File API: Blazor Web App Demos

[DevExpress Office & PDF File API](https://www.devexpress.com/products/net/office-file-api/) is a non-visual .NET library that allows you to generate, read, and modify PDF files and Microsoft Office documents (Word, Excel, PowerPoint)[^1].

[^1]: Office, Excel, Word, PowerPoint are trademarks or registered trademarks of Microsoft Corporation.

## What's Inside

This folder includes source code for DevExpress Office & PDF File API demos (Blazor Web Apps). If you used DevExpress Unified Installer, you can find the same code in your [local installation folder](https://docs.devexpress.com/OfficeFileAPI/15097/demo-application#find-the-source-code).

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

2. Open the following solution file in Visual Studio (see supported IDE and framework versions in [Prerequisites](https://docs.devexpress.com/OfficeFileAPI/405462/installation-guide/prerequisites)):

    * _Office File API/CS/Demos.Office.Web/Demos.Office.Web.sln_

3. Run the application.

## AI Setup (Requirements and Registration)

Review general requirements for AI-powered demos: [AI-powered Extensions for DevExpress Office File API](https://docs.devexpress.com/OfficeFileAPI/405645/ai-powered-extensions).

Note: AI services used for this demo have been rate limited. As such, you may experience performance-related delays when exploring the capabilities of DevExpress AI-powered Extensions.
When connected to your own AI model/service without rate limits, DevExpress AI-powered Extensions will perform seamlessly, without artificial delays.

In this project, AI configuration and service registration are defined in [Program.cs](./Program.cs):

- Read AI settings (`EndpointUrl`, `Key`, `DeploymentName`) from configuration (`AIIntegrationSettings`): [Program.cs#L32-L34](./Program.cs#L32-L34)
- Register chat client: [Program.cs#L47](./Program.cs#L47)
- Register DevExpress AI-powered document processing service: [Program.cs#L50-L51](./Program.cs#L50-L51)

## Project Structure (Demos.Office.Web)

Blazor Web App demo for DevExpress Office & PDF File API is organized as follows:

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
- [Services](./Services): shared helper/services used by demos (format conversion, extraction, utility logic).
- [DemoShell](./DemoShell): reusable layout, navigation, and search UI components for the demo host.
- [Documents](./Documents): source files and templates used in demo scenarios.
- [wwwroot](./wwwroot): static assets.

## Locate Code Related to DevExpress Office & PDF File API

If you want to find code directly related to DevExpress Office & PDF File API types and methods, start in the following locations:

1. Controller actions in [Controllers](./Controllers):

    This is the primary backend execution layer that implements document processing operations (load, modify, save, export). Check subfolders first (for example, [Controllers/PdfProcessing](./Controllers/PdfProcessing), [Controllers/WordProcessing](./Controllers/WordProcessing), [Controllers/SpreadsheetProcessing](./Controllers/SpreadsheetProcessing)).

2. Demo pages in [Pages](./Pages):

    These pages define user workflows, collect parameters, and invoke controller endpoints. Review corresponding subfolders to trace UI-to-backend flow.

3. Supporting models/services:

    - [Models](./Models): request/response shapes, file format descriptors, and demo metadata.
    - [Services](./Services): reusable helpers used across multiple demos.

4. Sample input files in [Documents](./Documents):

    Many demos use files from this folder as input templates or test data.

### Recommended Navigation Flow

1. Locate required page in [Pages](./Pages).
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
