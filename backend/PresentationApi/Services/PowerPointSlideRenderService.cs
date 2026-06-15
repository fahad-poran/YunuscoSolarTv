using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.Principal;
using Microsoft.Extensions.Options;
using PresentationApi.Options;

namespace PresentationApi.Services;

[SupportedOSPlatform("windows")]
public sealed class PowerPointSlideRenderService : ISlideRenderService
{
    private const int ComAccessDeniedHResult = unchecked((int)0x80070005);
    private const string PowerPointComCreationStage = "creating PowerPoint.Application COM instance";
    private const string PowerPointComAccessDeniedMessage =
        "PowerPoint COM access denied. The IIS app pool identity cannot start PowerPoint. Run the app pool under a dedicated Windows user, open PowerPoint once as that user, and grant DCOM Local Launch/Activation permission.";

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
        var stage = "initializing";

        try
        {
            stage = "locating PowerPoint COM registration";
            var powerPointType = Type.GetTypeFromProgID("PowerPoint.Application");
            if (powerPointType is null)
            {
                return Failure(
                    "Microsoft PowerPoint is not installed or is not registered for COM automation.",
                    pptxPath,
                    slidesDirectory,
                    stage);
            }

            stage = PowerPointComCreationStage;
            application = Activator.CreateInstance(powerPointType)!;

            stage = "configuring PowerPoint application";
            ConfigurePowerPointApplication(application);

            var absolutePptxPath = Path.GetFullPath(pptxPath);

            stage = "opening presentation";
            presentation = application.Presentations.Open(
                absolutePptxPath,
                0,  // msoFalse: read-write (Export fails when opened read-only)
                0,  // msoFalse: untitled
                -1); // msoTrue: with window

            stage = "exporting slides";
            List<string> normalizedNames = ExportSlides(presentation, slidesDirectory);

            if (normalizedNames.Count == 0)
            {
                return Failure(
                    "PowerPoint did not export any slide images.",
                    pptxPath,
                    slidesDirectory,
                    stage);
            }

            var exportedCount = normalizedNames.Count;
            _logger.LogInformation("Exported {Count} slides from {Path}", exportedCount, pptxPath);
            return SlideRenderResult.Succeeded(normalizedNames);
        }
        catch (COMException ex)
        {
            _logger.LogError(ex, "PowerPoint COM export failed for {Path}", pptxPath);
            return Failure(
                GetFailureMessage(stage, ex),
                pptxPath,
                slidesDirectory,
                stage,
                ex);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected slide export failure for {Path}", pptxPath);
            return Failure(
                GetFailureMessage(stage, ex),
                pptxPath,
                slidesDirectory,
                stage,
                ex);
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

    private SlideRenderResult Failure(
        string message,
        string pptxPath,
        string slidesDirectory,
        string stage,
        Exception? exception = null)
    {
        return SlideRenderResult.Failed(
            message,
            BuildDiagnostics(pptxPath, slidesDirectory, stage, exception));
    }

    private IReadOnlyDictionary<string, string?> BuildDiagnostics(
        string pptxPath,
        string slidesDirectory,
        string stage,
        Exception? exception)
    {
        var details = new Dictionary<string, string?>
        {
            ["stage"] = stage,
            ["pptxPath"] = SafeFullPath(pptxPath),
            ["pptxExists"] = File.Exists(pptxPath).ToString(),
            ["pptxSizeBytes"] = File.Exists(pptxPath) ? new FileInfo(pptxPath).Length.ToString() : null,
            ["slidesDirectory"] = SafeFullPath(slidesDirectory),
            ["slidesDirectoryExists"] = Directory.Exists(slidesDirectory).ToString(),
            ["processArchitecture"] = RuntimeInformation.ProcessArchitecture.ToString(),
            ["osArchitecture"] = RuntimeInformation.OSArchitecture.ToString(),
            ["osDescription"] = RuntimeInformation.OSDescription,
            ["currentUser"] = WindowsIdentity.GetCurrent().Name,
            ["userInteractive"] = Environment.UserInteractive.ToString()
        };

        if (exception is not null)
        {
            details["exceptionType"] = exception.GetType().FullName;
            details["exceptionMessage"] = exception.Message;
            details["exceptionHResult"] = $"0x{exception.HResult:X8}";
            details["innerExceptionType"] = exception.InnerException?.GetType().FullName;
            details["innerExceptionMessage"] = exception.InnerException?.Message;
            details["stackTrace"] = exception.StackTrace;
        }

        return details;
    }

    private static string SafeFullPath(string path)
    {
        try
        {
            return Path.GetFullPath(path);
        }
        catch
        {
            return path;
        }
    }

    private static string GetFailureMessage(string stage, Exception exception)
    {
        if (IsPowerPointComAccessDenied(stage, exception))
        {
            return PowerPointComAccessDeniedMessage;
        }

        return exception is COMException
            ? "PowerPoint could not export slides. Ensure PowerPoint is installed, activated, and available to the server account."
            : "Slide export failed unexpectedly.";
    }

    private static bool IsPowerPointComAccessDenied(string stage, Exception exception)
    {
        return stage == PowerPointComCreationStage &&
            exception.HResult == ComAccessDeniedHResult &&
            (exception is UnauthorizedAccessException || exception is COMException);
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
