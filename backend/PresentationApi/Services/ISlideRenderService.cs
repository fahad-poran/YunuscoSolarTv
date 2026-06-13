namespace PresentationApi.Services;

public interface ISlideRenderService
{
    Task<SlideRenderResult> RenderAsync(
        string pptxPath,
        string slidesDirectory,
        CancellationToken cancellationToken = default);
}
