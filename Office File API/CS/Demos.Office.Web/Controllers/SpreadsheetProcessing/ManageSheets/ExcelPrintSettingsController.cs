#region Code
using DevExpress.Spreadsheet;
using DevExpress.Office;
using DevExpress.Drawing.Printing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.ManageSheets {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelPrintSettingsController : SpreadsheetControllerBase {
        [HttpPost("print-settings/apply")]
        public async Task<IActionResult> Apply(IFormFile file, [FromForm] ExcelPrintSettingsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine output format
            var outputFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            var outputStream = ApplyPrintSettings(
                inputStream,
                new PageSettings {
                    PrintArea = "A1:I30",
                    PaperKind = model.PaperSize,
                    PageOrientation = model.PageOrientation,
                    CustomPaperWidth = model.CustomPaperWidth,
                    CustomPaperHeight = model.CustomPaperHeight
                },
                new Margins {
                    TopMargin = model.TopMargin,
                    BottomMargin = model.BottomMargin,
                    LeftMargin = model.LeftMargin,
                    RightMargin = model.RightMargin,
                    HeaderOffset = model.HeaderOffset,
                    FooterOffset = model.FooterOffset
                },
                new HeaderFooterSettings {
                    DifferentFirst = model.DifferentFirstPage,
                    HeaderLeftSectionText = model.HeaderLeftSectionText,
                    HeaderCenterSectionText = model.HeaderCenterSectionText,
                    HeaderRightSectionText = model.HeaderRightSectionText,
                    FooterLeftSectionText = model.FooterLeftSectionText,
                    FooterCenterSectionText = model.FooterCenterSectionText,
                    FooterRightSectionText = model.FooterRightSectionText
                },
                new PrintSettings {
                    FitToPage = model.FitToPage,
                    FitToWidth = model.FitToWidth,
                    FitToHeight = model.FitToHeight,
                    Scale = (int)System.Math.Round(model.Scale * 100, System.MidpointRounding.AwayFromZero),

                    PrintGridlines = model.PrintGridlines,
                    PrintHeadings = model.PrintHeadings,
                    BlackAndWhite = model.BlackAndWhite,

                    StartPage = model.StartPage,
                    StartType = WorksheetPageNumberingStartType.Restart,
                    PageOrder = model.PageOrder
                },
                outputFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream ApplyPrintSettings(
            Stream inputStream, PageSettings pageSettings, Margins margins,
            HeaderFooterSettings headerFooterSettings, PrintSettings printSettings, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Set measurement unit to inches
            workbook.Unit = DocumentUnit.Inch;

            var worksheet = workbook.Worksheets.First();

            // Get the active view
            var activeView = worksheet.ActiveView;

            // Specify print area
            if(!string.IsNullOrEmpty(pageSettings.PrintArea)) {
                var printRange = worksheet.Range[pageSettings.PrintArea];
                worksheet.SetPrintRange(printRange);
            }

            // Specify page settings
            activeView.Orientation = pageSettings.PageOrientation;

            if(pageSettings.PaperKind != DXPaperKind.Custom) {
                activeView.PaperKind = pageSettings.PaperKind;
            }
            else {
                activeView.SetCustomPaperSize(pageSettings.CustomPaperWidth, pageSettings.CustomPaperHeight);
            }

            // Specify margins
            activeView.Margins.Top = margins.TopMargin;
            activeView.Margins.Bottom = margins.BottomMargin;
            activeView.Margins.Left = margins.LeftMargin;
            activeView.Margins.Right = margins.RightMargin;
            activeView.Margins.Header = margins.HeaderOffset;
            activeView.Margins.Footer = margins.FooterOffset;

            // Specify header and footer settings
            var headerFooterOptions = worksheet.HeaderFooterOptions;
            headerFooterOptions.DifferentFirst = headerFooterSettings.DifferentFirst;

            if(headerFooterOptions.DifferentFirst) {
                headerFooterOptions.FirstHeader.Left = string.Format("{0}&BDevExpress", HeaderFooterCode.FontColor(4, -50));
                headerFooterOptions.FirstHeader.Right = "&A";

                headerFooterOptions.FirstFooter.Left = "&D";
                headerFooterOptions.FirstFooter.Right = string.Format("Page {0} of {1}", "&P", "&N");
            }

            headerFooterOptions.EvenHeader.Left = headerFooterOptions.OddHeader.Left = headerFooterSettings.HeaderLeftSectionText;
            headerFooterOptions.EvenHeader.Center = headerFooterOptions.OddHeader.Center = headerFooterSettings.HeaderCenterSectionText;
            headerFooterOptions.EvenHeader.Right = headerFooterOptions.OddHeader.Right = headerFooterSettings.HeaderRightSectionText;

            headerFooterOptions.EvenFooter.Left = headerFooterOptions.OddFooter.Left = headerFooterSettings.FooterLeftSectionText;
            headerFooterOptions.EvenFooter.Center = headerFooterOptions.OddFooter.Center = headerFooterSettings.FooterCenterSectionText;
            headerFooterOptions.EvenFooter.Right = headerFooterOptions.OddFooter.Right = headerFooterSettings.FooterRightSectionText;  

            // Specify print settings
            var printOptions = worksheet.PrintOptions;
            printOptions.FitToPage = printSettings.FitToPage;

            if(printOptions.FitToPage) {
                printOptions.FitToWidth = printSettings.FitToWidth;
                printOptions.FitToHeight = printSettings.FitToHeight;
            }
            else {
                printOptions.Scale = printSettings.Scale;
            }

            printOptions.PrintGridlines = printSettings.PrintGridlines;
            printOptions.PrintHeadings = printSettings.PrintHeadings;
            printOptions.BlackAndWhite = printSettings.BlackAndWhite;

            printOptions.PageNumbering.Start = printSettings.StartPage;
            printOptions.PageNumbering.StartType = printSettings.StartType;

            printOptions.PageOrder = printSettings.PageOrder;

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }

        public class PageSettings {
            public string PrintArea { get; set; } = string.Empty;
            public DXPaperKind PaperKind { get; set; } = DXPaperKind.Letter;
            public PageOrientation PageOrientation { get; set; } = PageOrientation.Default;
            public float CustomPaperWidth { get; set; } = 8.5f; // inches
            public float CustomPaperHeight { get; set; } = 11.0f; // inches
        }

        public class Margins {
            public float TopMargin { get; set; } = 1.0f;
            public float BottomMargin { get; set; } = 1.0f;
            public float LeftMargin { get; set; } = 1.0f;
            public float RightMargin { get; set; } = 1.0f;
            public float HeaderOffset { get; set; } = 0.5f;
            public float FooterOffset { get; set; } = 0.5f;
        }

        public class HeaderFooterSettings {
            public bool DifferentFirst { get; set; } = true;
            public string HeaderLeftSectionText { get; set; } = string.Empty;
            public string HeaderCenterSectionText { get; set; } = string.Empty;
            public string HeaderRightSectionText { get; set; } = string.Empty;
            public string FooterLeftSectionText { get; set; } = string.Empty;
            public string FooterCenterSectionText { get; set; } = string.Empty;
            public string FooterRightSectionText { get; set; } = string.Empty;
        }

        public class PrintSettings {
            public bool FitToPage { get; set; } = true;
            public int FitToWidth { get; set; } = 1;
            public int FitToHeight { get; set; } = 1;
            public int Scale { get; set; } = 100;

            public bool PrintGridlines { get; set; } = false;
            public bool PrintHeadings { get; set; } = false;
            public bool BlackAndWhite { get; set; } = false;

            public int StartPage { get; set; } = 1;
            public WorksheetPageNumberingStartType StartType { get; set; } = WorksheetPageNumberingStartType.Continuous;
            public PageOrder PageOrder { get; set; } = PageOrder.DownThenOver;
        }
        #endregion Code
    }
}
