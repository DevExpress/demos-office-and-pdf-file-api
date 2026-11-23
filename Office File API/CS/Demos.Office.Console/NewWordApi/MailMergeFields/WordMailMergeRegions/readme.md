# Word Master-Detail Mail Merge

This sample uses the **New DevExpress Word Processing Document API (WordDocument)** to generate an order report from a predefined master-detail template.

- Input: `Documents/MailMergeMasterDetail_Template.docx`.
- Data source: customers with orders and order details from `Documents/nwind_data.json`.
- `DefaultRecordExportMode`: `AllRecords` merges all customers. `ActiveRecord` merges only the record at `ActiveRecordIndex` (`0`).
- `DefaultDocumentExportMode`: `SingleFile` creates one document. `SeparateFiles` creates one document per record.
- `DefaultMergeMode`: `NewSection` starts each record in a new section.
- `OutputFormat`: `pdf`.
- Output: `Output/MailMerge_Result.pdf` beside the executable. With `SeparateFiles`, the sample writes `Output/record_<n>.pdf` files.