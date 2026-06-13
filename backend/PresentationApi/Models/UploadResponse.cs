namespace PresentationApi.Models;

public class UploadResponse
{
    public string Message { get; set; } = string.Empty;

    public string FileName { get; set; } = string.Empty;

    public long FileSizeBytes { get; set; }

    public DateTimeOffset UploadedAtUtc { get; set; }

    public int SlideCount { get; set; }

    public bool SlidesRendered { get; set; }

    public string? RenderWarning { get; set; }

    public IReadOnlyDictionary<string, string?>? RenderDiagnostics { get; set; }
}
