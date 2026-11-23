# Word Paragraph Formatting

This sample uses the new DevExpress Word Processing Document API (`WordDocument`) to format document paragraphs.

- Input file: `Documents/Sample.docx`
- Action: `ParagraphFormattingActionType.ApplyFormatting`. Set `ActionType` to `ClearFormatting` to reset attributes listed in `ResetAttributes` (all attributes by default).
- Range: `ParagraphFormattingRangeType.FirstParagraph`. Other values: `FirstSection`, `WholeDocument`.
- Formatting: left alignment, line spacing multiplier 1.15, left indent 0 inches, shading `#FFFFFF`, widow/orphan control on, keep with next off, keep lines together off
- Borders: `ApplyParagraphBorders` is `false`. When `true`, sample adds single 2 pt borders with color `#4682B4`.
- Output format: `docx` (`pdf`, `doc`, `dotx`, `rtf` are also supported)
- Output file: `Output/Formatted_Document.docx` beside executable