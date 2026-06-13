namespace PresentationApi.Options;

public class PresentationOptions
{
    public const string SectionName = "Presentation";

    public string StoragePath { get; set; } = "Files/current.pptx";

    public string SlidesPath { get; set; } = "Files/slides";

    public int SlideExportWidth { get; set; } = 1920;

    public int SlideExportHeight { get; set; } = 1080;

    public long MaxFileSizeBytes { get; set; } = 52_428_800; // 50 MB

    public string[] AllowedExtensions { get; set; } = [".pptx"];

    /// <summary>
    /// Public base URL of this API (e.g. https://api.example.com).
    /// Required for Office Online Viewer embed URLs.
    /// </summary>
    public string? PublicBaseUrl { get; set; }
}
