using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using PresentationApi.Models;
using PresentationApi.Options;
using PresentationApi.Services;

namespace PresentationApi.Controllers;

[ApiController]
[Route("api/presentation")]
public class PresentationController : ControllerBase
{
    private readonly PresentationOptions _options;
    private readonly IWebHostEnvironment _environment;
    private readonly ISlideRenderService _slideRenderService;
    private readonly ILogger<PresentationController> _logger;

    public PresentationController(
        IOptions<PresentationOptions> options,
        IWebHostEnvironment environment,
        ISlideRenderService slideRenderService,
        ILogger<PresentationController> logger)
    {
        _options = options.Value;
        _environment = environment;
        _slideRenderService = slideRenderService;
        _logger = logger;
    }

    [HttpPost("upload")]
    [RequestSizeLimit(52_428_800)]
    [RequestFormLimits(MultipartBodyLengthLimit = 52_428_800)]
    public async Task<ActionResult<UploadResponse>> Upload([FromForm] IFormFile? file, CancellationToken cancellationToken)
    {
        if (file is null || file.Length == 0)
        {
            return BadRequest(new ErrorResponse
            {
                Error = "No file uploaded.",
                Details = "Send a .pptx file using multipart/form-data with field name 'file'."
            });
        }

        if (file.Length > _options.MaxFileSizeBytes)
        {
            var maxMb = _options.MaxFileSizeBytes / (1024 * 1024);
            return BadRequest(new ErrorResponse
            {
                Error = "File is too large.",
                Details = $"Maximum allowed size is {maxMb} MB."
            });
        }

        var extension = Path.GetExtension(file.FileName);
        if (!_options.AllowedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            return BadRequest(new ErrorResponse
            {
                Error = "Invalid file type.",
                Details = "Only .pptx files are allowed."
            });
        }

        var storagePath = ResolveStoragePath();
        var directory = Path.GetDirectoryName(storagePath)!;
        Directory.CreateDirectory(directory);

        if (System.IO.File.Exists(storagePath))
        {
            System.IO.File.Delete(storagePath);
            _logger.LogInformation("Deleted existing presentation at {Path}", storagePath);
        }

        await using (var stream = new FileStream(storagePath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            await file.CopyToAsync(stream, cancellationToken);
        }

        var renderResult = await _slideRenderService.RenderAsync(
            storagePath,
            ResolveSlidesDirectory(),
            cancellationToken);

        var fileInfo = new FileInfo(storagePath);

        return Ok(new UploadResponse
        {
            Message = "Presentation uploaded successfully.",
            FileName = Path.GetFileName(storagePath),
            FileSizeBytes = fileInfo.Length,
            UploadedAtUtc = fileInfo.LastWriteTimeUtc,
            SlideCount = renderResult.Success ? renderResult.SlideFileNames.Count : 0,
            SlidesRendered = renderResult.Success,
            RenderWarning = renderResult.Success ? null : renderResult.ErrorMessage
        });
    }

    [HttpGet("view")]
    public ActionResult<PresentationViewResponse> ViewPresentation()
    {
        var storagePath = ResolveStoragePath();

        if (!System.IO.File.Exists(storagePath))
        {
            return Ok(new PresentationViewResponse { Exists = false });
        }

        var fileInfo = new FileInfo(storagePath);
        var fileUrl = BuildPublicUrl("api/presentation/file");
        var officeViewerUrl = fileUrl is null
            ? null
            : $"https://view.officeapps.live.com/op/embed.aspx?src={Uri.EscapeDataString(fileUrl)}";

        var slides = GetSlideInfos();

        return Ok(new PresentationViewResponse
        {
            Exists = true,
            FileName = fileInfo.Name,
            FileSizeBytes = fileInfo.Length,
            LastModifiedUtc = fileInfo.LastWriteTimeUtc,
            FileUrl = fileUrl,
            OfficeViewerUrl = officeViewerUrl,
            SlidesAvailable = slides.Count > 0,
            SlideCount = slides.Count,
            Slides = slides
        });
    }

    [HttpGet("file")]
    public IActionResult DownloadPresentation()
    {
        var storagePath = ResolveStoragePath();

        if (!System.IO.File.Exists(storagePath))
        {
            return NotFound(new ErrorResponse
            {
                Error = "No presentation found.",
                Details = "Upload a .pptx file before viewing."
            });
        }

        var stream = new FileStream(storagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        return File(
            stream,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "current.pptx");
    }

    [HttpGet("slides/{index:int}")]
    public IActionResult GetSlideImage(int index)
    {
        var slidesDirectory = ResolveSlidesDirectory();

        if (!Directory.Exists(slidesDirectory))
        {
            return NotFound(new ErrorResponse
            {
                Error = "Slides not found.",
                Details = "Re-upload the presentation to generate slide images."
            });
        }

        var slideFiles = Directory
            .GetFiles(slidesDirectory, "slide-*.png", SearchOption.TopDirectoryOnly)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (index < 1 || index > slideFiles.Count)
        {
            return NotFound(new ErrorResponse
            {
                Error = "Slide not found.",
                Details = $"Valid slide indexes are 1 through {slideFiles.Count}."
            });
        }

        var slidePath = slideFiles[index - 1];
        var stream = new FileStream(slidePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        return File(stream, "image/png");
    }

    private string ResolveStoragePath()
    {
        return Path.IsPathRooted(_options.StoragePath)
            ? _options.StoragePath
            : Path.Combine(_environment.ContentRootPath, _options.StoragePath);
    }

    private string ResolveSlidesDirectory()
    {
        return Path.IsPathRooted(_options.SlidesPath)
            ? _options.SlidesPath
            : Path.Combine(_environment.ContentRootPath, _options.SlidesPath);
    }

    private List<SlideInfo> GetSlideInfos()
    {
        var slidesDirectory = ResolveSlidesDirectory();

        if (!Directory.Exists(slidesDirectory))
        {
            return [];
        }

        return Directory
            .GetFiles(slidesDirectory, "slide-*.png", SearchOption.TopDirectoryOnly)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .Select((path, index) => new SlideInfo
            {
                Index = index + 1,
                Url = BuildPublicUrl($"api/presentation/slides/{index + 1}")
                    ?? $"{Request.Scheme}://{Request.Host}/api/presentation/slides/{index + 1}"
            })
            .ToList();
    }

    private string? BuildPublicUrl(string relativePath)
    {
        if (string.IsNullOrWhiteSpace(_options.PublicBaseUrl))
        {
            return null;
        }

        return $"{_options.PublicBaseUrl.TrimEnd('/')}/{relativePath.TrimStart('/')}";
    }
}
