namespace PresentationApi.Services;

public sealed class SlideRenderResult
{
    public bool Success { get; init; }

    public IReadOnlyList<string> SlideFileNames { get; init; } = [];

    public string? ErrorMessage { get; init; }

    public static SlideRenderResult Succeeded(IReadOnlyList<string> slideFileNames) =>
        new() { Success = true, SlideFileNames = slideFileNames };

    public static SlideRenderResult Failed(string errorMessage) =>
        new() { Success = false, ErrorMessage = errorMessage };
}
