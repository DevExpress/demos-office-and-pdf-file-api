# Word Format Sections

This sample uses the **New DevExpress Word Processing Document API (WordDocument)** to change page setup parameters for all document sections.

- Input: `Documents/Sample.docx`.
- `PaperSize`: `Letter`. `IsLandscape`: `false` (portrait). `CustomPaperWidth` / `CustomPaperHeight` (`8.5` / `11` inches) apply only when `PaperSize` is `Custom`.
- Margins: `TopMargin`, `BottomMargin`, `LeftMargin`, `RightMargin` are `1` inch. `MirroredMargins`: `false`.
- `HeaderOffset` / `FooterOffset`: `0.5` inch.
- `GutterSize`: `0` inches. `DefaultGutterPosition`: `Left` (also `Right` or `Top`).
- The sample converts inches to document units (300 per inch).
- `OutputFormat`: `docx`.
- Output: `Output/FormattedDocument.docx` beside the executable.