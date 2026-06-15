using System.Runtime.Versioning;
using PresentationApi.Options;
using PresentationApi.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.Configure<PresentationOptions>(
    builder.Configuration.GetSection(PresentationOptions.SectionName));

if (OperatingSystem.IsWindows())
{
    builder.Services.AddSingleton<ISlideRenderService, PowerPointSlideRenderService>();
}
else
{
    builder.Services.AddSingleton<ISlideRenderService, UnsupportedSlideRenderService>();
}

builder.Services.AddControllers();
builder.Services.AddOpenApi();

var allowedOrigins = builder.Configuration.GetSection("Cors:AllowedOrigins").Get<string[]>()
    ?? ["http://localhost:3000", "https://*.vercel.app"];

builder.Services.AddCors(options =>
{
    options.AddPolicy("Frontend", policy =>
    {
        policy.SetIsOriginAllowed(origin =>
            {
                if (string.IsNullOrWhiteSpace(origin))
                {
                    return false;
                }

                if (allowedOrigins.Contains(origin, StringComparer.OrdinalIgnoreCase))
                {
                    return true;
                }

                foreach (var pattern in allowedOrigins)
                {
                    if (pattern.Contains('*'))
                    {
                        var prefix = pattern.Split('*')[0];
                        if (origin.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }

                return false;
            })
            .AllowAnyHeader()
            .AllowAnyMethod();
    });
});

var app = builder.Build();

try
{
    var publicBaseUrl = app.Configuration["Presentation:PublicBaseUrl"];
    app.Logger.LogInformation("Server starting. PublicBaseUrl is: {url}", publicBaseUrl ?? "NOT CONFIGURED");
}
catch (Exception ex)
{
    Console.WriteLine($"Startup log failed: {ex.Message}");
}

app.UseCors("Frontend");

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

app.MapGet("/api/health", (IWebHostEnvironment environment) => Results.Ok(new
{
    status = "running",
    application = "PresentationApi",
    environment = environment.EnvironmentName,
    serverTimeUtc = DateTimeOffset.UtcNow
}));

// Only use HttpsRedirection if explicitly configured for HTTPS.
// For this LAN setup on port 75, we disable it to avoid 307 redirects that break CORS.
if (!app.Environment.IsDevelopment())
{
    // app.UseHttpsRedirection(); // Disabled to prevent redirect issues on LAN HTTP setup
}

app.MapControllers();

app.Run();
