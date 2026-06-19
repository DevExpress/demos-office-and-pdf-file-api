using DevExpress.Spreadsheet;

class Program {
    const string OutputFileName = "ThreadedComments_Document.xlsx";
    const string ThreadAuthor = "Sales Department";
    const string FirstCommentText = "The discount was approved by the Sales Director.";
    const string ReplyAuthor = "Jane Doe";
    const string ReplyCommentText = "Thank you for the confirmation!";
    const bool Resolved = true;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var outputStream = GenerateDocumentWithThreadedComments(ThreadAuthor, ReplyAuthor, FirstCommentText, ReplyCommentText, Resolved, OutputFormat);
            using var fileStream = File.Create(outputPath);
            outputStream.CopyTo(fileStream);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static Stream GenerateDocumentWithThreadedComments(string threadAuthor, string replyAuthor, string firstCommentText, string replyCommentText, bool resolved, DocumentFormat outputFormat) {
        using var workbook = new Workbook();
        var worksheet = workbook.Worksheets.First();

        worksheet.Cells["A1"].Value = "Select the 'B3' cell to read comments";
        worksheet.Cells["A3"].Value = "Discount:";

        var commentCell = worksheet.Cells["B3"];
        commentCell.Value = 0.15;
        commentCell.NumberFormat = "0%";
        worksheet.SelectedCell = commentCell;

        ThreadedComment threadedComment = worksheet.ThreadedComments.Add(commentCell, threadAuthor, firstCommentText);
        threadedComment.Resolved = resolved;
        threadedComment.Reply(replyAuthor, replyCommentText);

        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
