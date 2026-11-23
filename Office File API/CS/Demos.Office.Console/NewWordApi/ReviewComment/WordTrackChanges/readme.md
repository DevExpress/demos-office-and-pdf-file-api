# Word Track Changes

This sample uses the **New DevExpress Word Processing Document API (WordDocument)** to process revisions (tracked changes) in a Word document.

- Input: `Documents/TrackChanges_template.docx`.
- `RevisionAction`: `Accept` (default). The sample accepts all revisions. Set it to `Reject` to reject revisions, or to `None` to keep them as is.
- `AuthorFilter`: empty string, so revisions of all authors are processed. Set an author name to process only that author's revisions.
- `DisableTrackedChanges`: `false`, so change tracking stays enabled in the result.
- Output: `Output/TrackChanges_Result.docx` (DOCX) beside the executable.