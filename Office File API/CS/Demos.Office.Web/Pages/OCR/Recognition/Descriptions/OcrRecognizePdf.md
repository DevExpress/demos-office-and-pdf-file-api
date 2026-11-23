This module uses the **DevExpress OCR engine (`OcrEngine`)** and the **New DevExpress PDF Document API (`PdfDocument`)** to recognize text in PDF documents. You can process the predefined sample file (a scanned two-page document) or upload your own PDF.

Select the recognition mode: recognize rendered pages (works for any PDF) or only images embedded in pages (scanned documents). Specify the pages to process: all pages, the first page, or a custom page range. Use the **Detect page orientation** option to automatically detect and correct page rotation before recognition. Use the **Recognize and export to...** drop-down button to select the output format and download the result:

* **PDF** - a searchable PDF document that contains the processed pages and an invisible layer with the recognized text.
* **TXT** - recognized text as a plain text file.
* **JSON** - recognized text fragments with words, confidence values, and bounds (in PDF points).
