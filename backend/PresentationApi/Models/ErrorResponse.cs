namespace PresentationApi.Models;

public class ErrorResponse
{
    public string Error { get; set; } = string.Empty;

    public string? Details { get; set; }
}
