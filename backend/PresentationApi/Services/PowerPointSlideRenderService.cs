using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Extensions.Options;
using PresentationApi.Options;

namespace PresentationApi.Services;

[SupportedOSPlatform("windows")]
public sealed class PowerPointSlideRenderService : ISlideRenderService
{
    private readonly PresentationOptions _options;
    private readonly ILogger<PowerPointSlideRenderService> _logger;

    public PowerPointSlideRenderService(
        IOptions<PresentationOptions> options,
        ILogger<PowerPointSlideRenderService> logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    public Task<SlideRenderResult> RenderAsync(
        string pptxPath,
        string slidesDirectory,
        CancellationToken cancellationToken = default)
    {
        if (!OperatingSystem.IsWindows())
        {
            return Task.FromResult(SlideRenderResult.Failed("Slide rendering requires Windows with PowerPoint installed."));
        }

        var tcs = new TaskCompletionSource<SlideRenderResult>();

        var thread = new Thread(() =>
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                var result = RenderOnStaThread(pptxPath, slidesDirectory);
                tcs.SetResult(result);
            }
            catch (OperationCanceledException)
            {
                tcs.SetCanceled(cancellationToken);
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();

        return tcs.Task;
    }

    private SlideRenderResult RenderOnStaThread(string pptxPath, string slidesDirectory)
    {
        if (!File.Exists(pptxPath))
        {
            return SlideRenderResult.Failed("Presentation file was not found.");
        }

        Directory.CreateDirectory(slidesDirectory);
        ClearDirectory(slidesDirectory);

        dynamic? application = null;
        dynamic? presentation = null;

        try
        {
            var powerPointType = Type.GetTypeFromProgID("PowerPoint.Application");
            if (powerPointType is null)
            {
                return SlideRenderResult.Failed("Microsoft PowerPoint is not installed on this machine.");
            }

            application = Activator.CreateInstance(powerPointType)!;
            ConfigurePowerPointApplication(application);

            var absolutePptxPath = Path.GetFullPath(pptxPath);

            presentation = application.Presentations.Open(
                absolutePptxPath,
                0,  // msoFalse: read-write (Export fails when opened read-only)
                0,  // msoFalse: untitled
                -1); // msoTrue: with window

            List<string> normalizedNames = ExportSlides(presentation, slidesDirectory);

            if (normalizedNames.Count == 0)
            {
                return SlideRenderResult.Failed("PowerPoint did not export any slide images.");
            }

            var exportedCount = normalizedNames.Count;
            _logger.LogInformation("Exported {Count} slides from {Path}", exportedCount, pptxPath);
            return SlideRenderResult.Succeeded(normalizedNames);
        }
        catch (COMException ex)
        {
            _logger.LogError(ex, "PowerPoint COM export failed for {Path}", pptxPath);
            return SlideRenderResult.Failed("PowerPoint could not export slides. Ensure PowerPoint is installed and licensed.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected slide export failure for {Path}", pptxPath);
            return SlideRenderResult.Failed("Slide export failed unexpectedly.");
        }
        finally
        {
            if (presentation is not null)
            {
                try
                {
                    presentation.Close();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to close PowerPoint presentation.");
                }

                Marshal.FinalReleaseComObject(presentation);
            }

            if (application is not null)
            {
                try
                {
                    application.Quit();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to quit PowerPoint.");
                }

                Marshal.FinalReleaseComObject(application);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private List<string> ExportSlides(dynamic presentation, string slidesDirectory)
    {
        var normalizedNames = new List<string>();
        var slideCount = (int)presentation.Slides.Count;

        for (var index = 1; index <= slideCount; index++)
        {
            dynamic slide = presentation.Slides[index];
            var normalizedName = $"slide-{index:D3}.png";
            var destinationPath = Path.GetFullPath(Path.Combine(slidesDirectory, normalizedName));

            slide.Export(
                destinationPath,
                "PNG",
                _options.SlideExportWidth,
                _options.SlideExportHeight);

            normalizedNames.Add(normalizedName);
        }

        return normalizedNames;
    }

    private void ConfigurePowerPointApplication(dynamic application)
    {
        try
        {
            application.DisplayAlerts = 1; // ppAlertsNone
        }
        catch (COMException ex)
        {
            _logger.LogWarning(ex, "Could not suppress PowerPoint alerts.");
        }
    }

    private static void ClearDirectory(string directory)
    {
        if (!Directory.Exists(directory))
        {
            return;
        }

        foreach (var file in Directory.GetFiles(directory))
        {
            File.Delete(file);
        }
    }
}
