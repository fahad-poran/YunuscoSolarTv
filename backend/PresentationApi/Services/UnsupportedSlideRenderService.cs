namespace PresentationApi.Services;

public sealed class UnsupportedSlideRenderService : ISlideRenderService
{
    public Task<SlideRenderResult> RenderAsync(
        string pptxPath,
        string slidesDirectory,
        CancellationToken cancellationToken = default)
    {
        return Task.FromResult(
            SlideRenderResult.Failed("Slide rendering is only supported on Windows with PowerPoint installed."));
    }
}
