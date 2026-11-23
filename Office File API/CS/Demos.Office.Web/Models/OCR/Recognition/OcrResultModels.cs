namespace BlazorDemo.OfficeFileApi.Models.OCR.Recognition {
    public record OcrBounds(float X, float Y, float Width, float Height);

    public record OcrWord(string Text, float? Confidence, OcrBounds Bounds);

    public record OcrFragment(string Text, float? Confidence, OcrBounds Bounds, IReadOnlyList<OcrWord> Words);

    public record OcrPage(int Page, IReadOnlyList<OcrFragment> Fragments);
}
