#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.MergeSplit;

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.MergeSplit {
    [ApiController]
    [Route("api/wordprocessing")]
    public class MergeController : WordControllerBase {
        [HttpPost("word-merge/process")]
        public async Task<IActionResult> MergeWordDocuments(List<IFormFile> files, [FromForm] WordMergeDocumentModel model) {
            var inputStreams = new List<Stream>();
            foreach(var file in files) {
                var ms = new MemoryStream();
                await file.CopyToAsync(ms);
                ms.Position = 0;
                inputStreams.Add(ms);
            }

            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);
            var outputStream = MergeDocuments(inputStreams, targetFormat, model.IsPdfOutput);
            outputStream.Position = 0;
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream MergeDocuments(List<Stream> inputStreams, DocumentFormat targetFormat, bool pdfOutput) {
            using var sourceWordProcessor = new RichEditDocumentServer();
            using var targetWordProcessor = new RichEditDocumentServer();
            var sourceDocument = sourceWordProcessor.Document;
            var targetDocument = targetWordProcessor.Document;

            for(int i = 0; i < inputStreams.Count; i++) {
                sourceWordProcessor.LoadDocument(inputStreams[i]);

                // Unlink all headers and footers
                var targetSection = targetDocument.Sections[targetDocument.Sections.Count - 1];
                targetSection.UnlinkHeaderFromPrevious();
                targetSection.UnlinkFooterFromPrevious();

                targetSection.UnlinkHeaderFromPrevious(HeaderFooterType.First);
                targetSection.UnlinkFooterFromPrevious(HeaderFooterType.First);

                targetSection.UnlinkHeaderFromPrevious(HeaderFooterType.Even);
                targetSection.UnlinkFooterFromPrevious(HeaderFooterType.Even);

                // Append each document section one-by-one
                AppendDocumentSection(sourceDocument, targetDocument);

                if(i < inputStreams.Count - 1) {
                    // Append a new section that starts
                    // on the new page
                    targetDocument.AppendSection();
                    targetDocument.Sections.Last().StartType = SectionStartType.NextPage;
                }
            }
            var outputStream = new MemoryStream();
            if(pdfOutput)
                targetWordProcessor.ExportToPdf(outputStream);
            else
                targetWordProcessor.SaveDocument(outputStream, targetFormat);

            outputStream.Position = 0;
            return outputStream;
        }

        void AppendDocumentSection(Document source, Document target) {
            var lastSectionIndexBeforeAppending = target.Sections.Count - 1;
            var sourceSectionCount = source.Sections.Count;

            // Append document body
            var appendedRange = target.AppendDocumentContent(source.Range);
            target.Delete(target.Paragraphs.Get(target.CreatePosition(appendedRange.Start.ToInt() - 1)).Range);

            target.DifferentOddAndEvenPages = source.DifferentOddAndEvenPages;

            for(int i = 0; i < sourceSectionCount; i++) {
                var sourceSection = source.Sections[i];
                var targetSection = target.Sections[lastSectionIndexBeforeAppending + i];

                // Copy headers and footers
                AppendHeader(sourceSection, targetSection, HeaderFooterType.Odd);
                AppendFooter(sourceSection, targetSection, HeaderFooterType.Odd);

                AppendHeader(sourceSection, targetSection, HeaderFooterType.Even);
                AppendFooter(sourceSection, targetSection, HeaderFooterType.Even);

                AppendHeader(sourceSection, targetSection, HeaderFooterType.First);
                AppendFooter(sourceSection, targetSection, HeaderFooterType.First);
            }
        }

        void AppendHeader(Section sourceSection, Section targetSection, HeaderFooterType headerFooterType) {
            // Clear headers in the target document
            if(!sourceSection.HasHeader(headerFooterType)) {
                if(targetSection.HasHeader(headerFooterType)) {
                    var targetHeader = targetSection.BeginUpdateHeader(headerFooterType);
                    targetHeader.Delete(targetHeader.Range);
                    targetSection.EndUpdateHeader(targetHeader);
                }
                return;
            }

            // Insert header content
            var source = sourceSection.BeginUpdateHeader(headerFooterType);
            var target = targetSection.BeginUpdateHeader(headerFooterType);
            target.Delete(target.Range);
            target.InsertDocumentContent(target.Range.Start, source.Range, InsertOptions.KeepSourceFormatting);

            // Delete empty paragraphs
            var emptyParagraph = target.CreateRange(target.Range.End.ToInt() - 1, 1);
            target.Delete(emptyParagraph);

            sourceSection.EndUpdateHeader(source);
            targetSection.EndUpdateHeader(target);
        }

        void AppendFooter(Section sourceSection, Section targetSection, HeaderFooterType headerFooterType) {
            // Clear footers in the target document
            if(!sourceSection.HasFooter(headerFooterType)) {
                if(targetSection.HasFooter(headerFooterType)) {
                    var targetFooter = targetSection.BeginUpdateFooter(headerFooterType);
                    targetFooter.Delete(targetFooter.Range);
                    targetSection.EndUpdateFooter(targetFooter);
                }
                return;
            }

            // Insert footer content
            var source = sourceSection.BeginUpdateFooter(headerFooterType);
            var target = targetSection.BeginUpdateFooter(headerFooterType);
            target.Delete(target.Range);
            target.InsertDocumentContent(target.Range.Start, source.Range, InsertOptions.KeepSourceFormatting);

            // Delete empty paragraphs
            var emptyParagraph = target.CreateRange(target.Range.End.ToInt() - 1, 1);
            target.Delete(emptyParagraph);

            sourceSection.EndUpdateFooter(source);
            targetSection.EndUpdateFooter(target);
        }
        #endregion Code
    }
}
