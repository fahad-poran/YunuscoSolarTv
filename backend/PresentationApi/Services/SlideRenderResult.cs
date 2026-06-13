namespace PresentationApi.Services;

public sealed class SlideRenderResult
{
    public bool Success { get; init; }

    public IReadOnlyList<string> SlideFileNames { get; init; } = [];

    public string? ErrorMessage { get; init; }

    public IReadOnlyDictionary<string, string?> ErrorDetails { get; init; } =
        new Dictionary<string, string?>();

    public static SlideRenderResult Succeeded(IReadOnlyList<string> slideFileNames) =>
        new() { Success = true, SlideFileNames = slideFileNames };

    public static SlideRenderResult Failed(
        string errorMessage,
        IReadOnlyDictionary<string, string?>? errorDetails = null) =>
        new()
        {
            Success = false,
            ErrorMessage = errorMessage,
            ErrorDetails = errorDetails ?? new Dictionary<string, string?>()
        };
}
