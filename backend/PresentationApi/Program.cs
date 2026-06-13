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

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

app.UseCors("Frontend");

if (!app.Environment.IsDevelopment())
{
    app.UseHttpsRedirection();
}

app.MapControllers();

app.Run();
