namespace PresentationApi.Models;

public class PresentationViewResponse
{
    public bool Exists { get; set; }

    public string? FileName { get; set; }

    public long? FileSizeBytes { get; set; }

    public DateTimeOffset? LastModifiedUtc { get; set; }

    /// <summary>
    /// Direct URL to download/stream the .pptx file.
    /// </summary>
    public string? FileUrl { get; set; }

    /// <summary>
    /// Office Online Viewer embed URL (only when PublicBaseUrl is configured).
    /// </summary>
    public string? OfficeViewerUrl { get; set; }

    public bool SlidesAvailable { get; set; }

    public int SlideCount { get; set; }

    public List<SlideInfo> Slides { get; set; } = [];
}
