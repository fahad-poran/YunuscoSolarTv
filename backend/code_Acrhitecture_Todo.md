# 🏗️ Backend Code Architecture Guide (Junior Developer)

> 📌 **Project:** YunuscoSolarTV — Presentation API  
> 👤 **Audience:** Junior .NET developers who want to understand, maintain, or rebuild this backend **without AI assistance**  
> 📂 **Scope:** Backend only (`backend/PresentationApi/`)

---

## 📑 Table of contents

| # | Section |
|---|---------|
| 1 | [What this backend does](#1-what-this-backend-does) |
| 2 | [Packages and tools used](#2-packages-and-tools-used) |
| 3 | [Build it yourself — step by step](#3-build-it-yourself--step-by-step) |
| 4 | [Folder structure explained](#4-folder-structure-explained) |
| 5 | [How a request flows through the code](#5-how-a-request-flows-through-the-code) |
| 6 | [File-by-file code explanation](#6-file-by-file-code-explanation) |
| 7 | [Important concepts you must understand](#7-important-concepts-you-must-understand) |
| 8 | [Common problems and how to debug](#8-common-problems-and-how-to-debug) |
| 9 | [Junior developer TODO checklist](#9-junior-developer-todo-checklist) |
| 10 | [Publish to IIS (Windows server)](#10-publish-to-iis-windows-server) |

---

### 🧭 Quick navigation by topic

| I want to… | Go to |
|------------|-------|
| Understand endpoints | [§1](#1-what-this-backend-does) · [Quick reference](#-quick-reference--endpoints) |
| Build from scratch | [§3](#3-build-it-yourself--step-by-step) |
| Deploy to IIS | [§10](#10-publish-to-iis-windows-server) |
| Fix CORS / upload errors | [§8](#8-common-problems-and-how-to-debug) · [§10.6](#106-cors--browser-errors-we-fixed) |
| Understand PowerPoint export | [§6.3](#63-powerpointsliderenderservicecs) · [§10.7](#107-powerpoint--slide-export-on-the-server) |

---

## 1. What this backend does

This API has **one main job**:

| Step | Action |
|------|--------|
| 1️⃣ | Accept a `.pptx` file upload |
| 2️⃣ | Save it to disk (always the same file: `Files/current.pptx`) |
| 3️⃣ | Use **Microsoft PowerPoint** (installed on the server) to export each slide as a PNG image |
| 4️⃣ | Let the frontend fetch slide images and metadata |

```
🌐 Client (browser)
    │
    ├─ POST /api/presentation/upload   → 💾 save pptx + 🖼️ export PNG slides
    ├─ GET  /api/presentation/view     → 📋 JSON metadata + slide URLs
    ├─ GET  /api/presentation/file     → 📎 raw .pptx download
    └─ GET  /api/presentation/slides/1 → 🖼️ single slide PNG
```

> 💡 **Why PNG export instead of sending the .pptx to the browser?**  
> Browsers cannot render PowerPoint files with correct fonts and layout. Exporting slides as images using PowerPoint itself gives pixel-perfect results on TV screens.

---

## 2. Packages and tools used

### 2.1 ⚙️ Runtime & SDK

| Tool | Version | Purpose |
|------|---------|---------|
| .NET SDK | 10.x | Build and run the API |
| Target framework | `net10.0` | Project runtime |

### 2.2 📦 NuGet packages

Only **one** external NuGet package is used:

| Package | Version | Why we use it |
|---------|---------|---------------|
| `Microsoft.AspNetCore.OpenApi` | 10.0.8 | Generates OpenAPI/Swagger metadata in Development so you can inspect endpoints |

**No extra packages** are needed for:
- File upload (built into ASP.NET Core)
- CORS (built into ASP.NET Core)
- PowerPoint export (uses Windows COM — built into .NET + Windows)

### 2.3 🪟 Windows dependencies (not NuGet)

| Dependency | Why |
|------------|-----|
| **Windows OS** | PowerPoint COM automation only works on Windows |
| **Microsoft PowerPoint** | Installed desktop app used to export slide PNGs |
| **COM Interop** | .NET talks to PowerPoint through Windows Component Object Model |

> **Junior dev note:** COM is how Windows programs talk to other Windows programs (like Excel, Word, PowerPoint). We use `dynamic` + `Type.GetTypeFromProgID("PowerPoint.Application")` instead of adding a NuGet interop package.

---

## 3. 🔨 Build it yourself — step by step

Follow these steps in order if you want to recreate this backend manually.

### Step 1 — Create the Web API project

Open PowerShell / terminal:

```powershell
mkdir YunuscoSolarTV\backend
cd YunuscoSolarTV\backend
dotnet new webapi -n PresentationApi --use-controllers
cd PresentationApi
```

**Why `--use-controllers`?**  
We want traditional Controller classes (easier for juniors to read than Minimal APIs for this project).

---

### Step 2 — Add the NuGet package

```powershell
dotnet add package Microsoft.AspNetCore.OpenApi
```

---

### Step 3 — Restrict to Windows (optional but recommended)

Edit `PresentationApi.csproj`:

```xml
<SupportedOSPlatform>windows</SupportedOSPlatform>
```

**Why?** Slide export only works on Windows. This documents the platform requirement at compile time.

---

### Step 4 — Create folders

```powershell
mkdir Controllers, Models, Options, Services
```

| Folder | Responsibility |
|--------|----------------|
| `Controllers/` | HTTP endpoints (upload, view, download) |
| `Models/` | JSON response shapes returned to frontend |
| `Options/` | Settings loaded from `appsettings.json` |
| `Services/` | Business logic (PowerPoint slide export) |

This is called **separation of concerns** — each folder has one job.

---

### Step 5 — Create configuration (`Options/PresentationOptions.cs`)

Map settings from `appsettings.json` to a C# class.

**Why a separate Options class?**
- Avoids magic strings like `"Files/current.pptx"` scattered in code
- Easy to change paths/sizes without recompiling logic
- ASP.NET Core binds JSON config to this class automatically

Add to `appsettings.json`:

```json
"Presentation": {
  "StoragePath": "Files/current.pptx",
  "SlidesPath": "Files/slides",
  "SlideExportWidth": 1920,
  "SlideExportHeight": 1080,
  "MaxFileSizeBytes": 52428800,
  "AllowedExtensions": [ ".pptx" ],
  "PublicBaseUrl": ""
}
```

---

### Step 6 — Create response models (`Models/`)

Create plain C# classes for API responses:

- `UploadResponse.cs` — returned after upload
- `PresentationViewResponse.cs` — returned by `/view`
- `SlideInfo.cs` — one slide URL + index
- `ErrorResponse.cs` — consistent error JSON

**Why separate model classes?**
- Controller returns structured JSON, not anonymous objects
- Frontend knows exactly what fields to expect
- Easier to maintain and document

---

### Step 7 — Create slide render service (`Services/`)

Create three files:

1. `ISlideRenderService.cs` — interface (contract)
2. `SlideRenderResult.cs` — success/failure result object
3. `PowerPointSlideRenderService.cs` — real Windows implementation
4. `UnsupportedSlideRenderService.cs` — fallback for non-Windows

**Why an interface (`ISlideRenderService`)?**
- Controller does not need to know *how* slides are rendered
- You can swap implementations later (e.g. LibreOffice) without changing the controller
- This is **Dependency Injection (DI)** — a core .NET pattern

---

### Step 8 — Create the controller (`Controllers/PresentationController.cs`)

Add 4 endpoints:

| Method | Route | Action |
|--------|-------|--------|
| POST | `/api/presentation/upload` | Save file + trigger slide export |
| GET | `/api/presentation/view` | Return metadata + slide list |
| GET | `/api/presentation/file` | Stream the `.pptx` |
| GET | `/api/presentation/slides/{index}` | Stream one PNG slide |

**Why one controller?**  
All endpoints relate to the same resource: "the current presentation".

---

### Step 9 — Wire everything in `Program.cs`

`Program.cs` is the application entry point. It must:

1. Load config → `PresentationOptions`
2. Register services → `ISlideRenderService`
3. Enable CORS → allow frontend on Vercel/localhost
4. Map controllers

**Why CORS?**  
Frontend runs on a different URL (e.g. Vercel). Browsers block cross-origin requests unless the API explicitly allows them.

---

### Step 10 — Run and test

```powershell
dotnet run
```

Test with browser or curl:

```powershell
# Check metadata (before upload)
Invoke-RestMethod http://localhost:5025/api/presentation/view

# Upload (replace path with your pptx)
curl -F "file=@C:\path\to\test.pptx" http://localhost:5025/api/presentation/upload
```

**Requirement:** PowerPoint must be installed on the machine running the API.

---

## 4. 📁 Folder structure explained

```
backend/PresentationApi/
│
├── Program.cs                          ← App startup, DI, CORS, middleware
├── PresentationApi.csproj                ← Project file + NuGet references
├── appsettings.json                      ← Configuration values
│
├── Controllers/
│   └── PresentationController.cs         ← HTTP endpoints only
│
├── Services/
│   ├── ISlideRenderService.cs            ← Interface
│   ├── SlideRenderResult.cs              ← Result object
│   ├── PowerPointSlideRenderService.cs   ← PowerPoint COM export (Windows)
│   └── UnsupportedSlideRenderService.cs    ← Safe fallback (non-Windows)
│
├── Options/
│   └── PresentationOptions.cs            ← Strongly-typed config
│
├── Models/
│   ├── UploadResponse.cs
│   ├── PresentationViewResponse.cs
│   ├── SlideInfo.cs
│   └── ErrorResponse.cs
│
├── Files/                                ← Created at runtime (not in source control)
│   ├── current.pptx                      ← Latest uploaded presentation
│   └── slides/
│       ├── slide-001.png
│       ├── slide-002.png
│       └── ...
│
└── Properties/
    └── launchSettings.json               ← Local dev URL (port 5025)
```

---

## 5. 🔀 How a request flows through the code

### Upload flow (`POST /api/presentation/upload`)

```
1. Browser sends multipart/form-data with field name "file"
2. ASP.NET Core routes to PresentationController.Upload()
3. Controller validates:
   - file exists?
   - size <= 50 MB?
   - extension is .pptx?
4. Controller deletes old current.pptx (replace logic)
5. Controller saves new file to Files/current.pptx
6. Controller calls ISlideRenderService.RenderAsync()
7. PowerPointSlideRenderService:
   - starts STA thread (required for COM)
   - opens PowerPoint via COM
   - exports each slide to Files/slides/slide-XXX.png
   - closes PowerPoint
8. Controller returns UploadResponse JSON
```

### View flow (`GET /api/presentation/view`)

```
1. Controller checks if Files/current.pptx exists
2. If not → return { exists: false }
3. If yes → scan Files/slides/ for PNG files
4. Build list of slide URLs
5. Return PresentationViewResponse JSON
```

### Slide image flow (`GET /api/presentation/slides/3`)

```
1. Controller lists PNG files in Files/slides/
2. Picks file at index 3 (1-based)
3. Returns image/png stream to browser
```

---

## 6. 📄 File-by-file code explanation

### 6.1 `Program.cs`

```csharp
builder.Services.Configure<PresentationOptions>(
    builder.Configuration.GetSection(PresentationOptions.SectionName));
```

**Why:** Reads `"Presentation"` section from `appsettings.json` into `PresentationOptions` class.

---

```csharp
if (OperatingSystem.IsWindows())
{
    builder.Services.AddSingleton<ISlideRenderService, PowerPointSlideRenderService>();
}
else
{
    builder.Services.AddSingleton<ISlideRenderService, UnsupportedSlideRenderService>();
}
```

**Why:**
- `Singleton` = one instance for the whole app lifetime (fine for stateless export service)
- Windows gets real PowerPoint exporter
- Linux/Mac gets safe fallback that returns a clear error message

---

```csharp
builder.Services.AddCors(options => { ... });
```

**Why:** Frontend is hosted separately (Vercel). Without CORS, browser blocks API calls.

**Wildcard support (`https://*.vercel.app`):**  
We manually check if origin starts with `https://` before the `*` so any Vercel deployment works.

---

```csharp
if (!app.Environment.IsDevelopment())
{
    app.UseHttpsRedirection();
}
```

**Why:** In local dev we run HTTP only (port 5025). Forcing HTTPS redirect in dev caused the warning:  
`Failed to determine the https port for redirect.`

---

### 6.2 `PresentationController.cs`

#### Constructor injection

```csharp
public PresentationController(
    IOptions<PresentationOptions> options,
    IWebHostEnvironment environment,
    ISlideRenderService slideRenderService,
    ILogger<PresentationController> logger)
```

**Why:** ASP.NET Core **Dependency Injection** automatically provides these objects. You don't write `new PowerPointSlideRenderService()` in the controller.

| Injected service | Used for |
|------------------|----------|
| `IOptions<PresentationOptions>` | Read config (paths, size limits) |
| `IWebHostEnvironment` | Get app root folder path |
| `ISlideRenderService` | Export slides after upload |
| `ILogger` | Write logs to console / file |

---

#### Upload validation

```csharp
[HttpPost("upload")]
[RequestSizeLimit(52_428_800)]
[RequestFormLimits(MultipartBodyLengthLimit = 52_428_800)]
public async Task<ActionResult<UploadResponse>> Upload([FromForm] IFormFile? file, ...)
```

**Why each attribute:**

| Attribute | Purpose |
|-----------|---------|
| `[HttpPost("upload")]` | Maps to `POST /api/presentation/upload` |
| `[RequestSizeLimit]` | Rejects requests larger than 50 MB at server level |
| `[RequestFormLimits]` | Same limit specifically for multipart form uploads |
| `[FromForm]` | Tells ASP.NET to read file from HTML form field |
| `IFormFile` | Built-in type for uploaded files |

---

#### Replace logic (delete old file first)

```csharp
if (System.IO.File.Exists(storagePath))
{
    System.IO.File.Delete(storagePath);
}
```

**Why:** Business requirement — only **one** presentation exists at a time (`current.pptx`). TV display always shows the latest upload.

---

#### Save file

```csharp
await using (var stream = new FileStream(storagePath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
{
    await file.CopyToAsync(stream, cancellationToken);
}
```

**Why:**
- `FileMode.CreateNew` — fails if file somehow still exists (extra safety)
- `FileShare.None` — no other process can read/write while saving
- `await using` — stream is disposed automatically (prevents file lock bugs)
- `CopyToAsync` — efficient for large files, supports cancellation

---

#### Path resolution helpers

```csharp
private string ResolveStoragePath()
{
    return Path.IsPathRooted(_options.StoragePath)
        ? _options.StoragePath
        : Path.Combine(_environment.ContentRootPath, _options.StoragePath);
}
```

**Why:**
- `Files/current.pptx` is a **relative** path → combine with app root
- Supports absolute paths too (useful in production if you mount a network drive)
- Same pattern used for `SlidesPath`

---

#### Slide URL building

```csharp
Url = BuildPublicUrl($"api/presentation/slides/{index + 1}")
    ?? $"{Request.Scheme}://{Request.Host}/api/presentation/slides/{index + 1}"
```

**Why:**
- In **production**, set `PublicBaseUrl` in config so URLs are correct behind reverse proxies
- In **local dev**, fall back to `http://localhost:5025/...` from the current request

---

### 6.3 `PowerPointSlideRenderService.cs`

This is the most complex file. Read it carefully.

#### Why a separate STA thread?

```csharp
var thread = new Thread(() => { ... });
thread.SetApartmentState(ApartmentState.STA);
thread.Start();
```

**Why:**  
PowerPoint COM objects require an **STA (Single-Threaded Apartment)** thread. ASP.NET Core request threads are MTA by default. Without STA, COM calls fail or behave randomly.

**Pattern used:**
1. Create `TaskCompletionSource<SlideRenderResult>`
2. Run COM work on STA background thread
3. Return `tcs.Task` to caller (controller still uses `async/await`)

---

#### Why `Type.GetTypeFromProgID` instead of a NuGet package?

```csharp
var powerPointType = Type.GetTypeFromProgID("PowerPoint.Application");
application = Activator.CreateInstance(powerPointType)!;
```

**Why:**
- Finds installed PowerPoint on the machine
- No need for `Microsoft.Office.Interop.PowerPoint` NuGet package
- Uses `dynamic` to call methods like `Presentations.Open()` without compile-time interop references

**Trade-off:** No IntelliSense on PowerPoint methods — you must know the API from Microsoft docs.

---

#### Why open read-write (not read-only)?

```csharp
presentation = application.Presentations.Open(
    absolutePptxPath,
    0,  // read-write — Export FAILS if read-only
    0,
    -1);
```

**Why:** We learned this from a real bug. Opening as read-only caused:  
`Presentation.Export : Invalid request. Presentation cannot be modified.`

Export is treated as a modification operation by PowerPoint.

---

#### Why NOT set `Application.Visible = false`?

We intentionally removed hiding PowerPoint because some installs throw:  
`Hiding the application window is not allowed.`

Only `DisplayAlerts` is suppressed (to avoid popup dialogs blocking automation).

---

#### Per-slide export (not whole presentation export)

```csharp
for (var index = 1; index <= slideCount; index++)
{
    dynamic slide = presentation.Slides[index];
    slide.Export(destinationPath, "PNG", width, height);
}
```

**Why:**
- PowerPoint slide indexes are **1-based** (not 0-based like C# arrays)
- Direct control over output filenames: `slide-001.png`, `slide-002.png`
- Avoids renaming `Slide1.PNG`, `Slide2.PNG` after batch export

---

#### COM cleanup in `finally`

```csharp
presentation.Close();
application.Quit();
Marshal.FinalReleaseComObject(presentation);
Marshal.FinalReleaseComObject(application);
GC.Collect();
GC.WaitForPendingFinalizers();
```

**Why:**  
COM objects are unmanaged memory. If you don't close PowerPoint:
- `POWERPNT.EXE` stays running in Task Manager
- Next upload may fail or open wrong instance
- Server runs out of resources over time

**Always clean up COM objects in `finally`.**

---

### 6.4 `UnsupportedSlideRenderService.cs`

```csharp
return Task.FromResult(
    SlideRenderResult.Failed("Slide rendering is only supported on Windows..."));
```

**Why:**  
If someone runs `dotnet run` on Linux/Mac, the app still starts. Upload works, but slide export returns a clear warning instead of crashing.

---

### 6.5 Model classes — why so simple?

Example `UploadResponse`:

```csharp
public class UploadResponse
{
    public string Message { get; set; } = string.Empty;
    public int SlideCount { get; set; }
    public bool SlidesRendered { get; set; }
    public string? RenderWarning { get; set; }
}
```

**Why:**
- ASP.NET Core serializes public properties to JSON automatically
- `RenderWarning` tells frontend if PowerPoint export failed but upload succeeded
- Nullable `string?` means field can be `null` in JSON

---

## 7. 💡 Important concepts you must understand

### 7.1 Dependency Injection (DI)

**What:** Instead of creating objects yourself, you declare what you need in the constructor and the framework provides them.

**In this project:**
```
Controller → needs → ISlideRenderService
Program.cs → registers → PowerPointSlideRenderService as ISlideRenderService
```

**Practice exercise:** Try injecting `ILogger<PowerPointSlideRenderService>` and add more log messages.

---

### 7.2 Options pattern

**What:** Bind `appsettings.json` to a strongly-typed C# class.

**Benefits:**
- Change file paths without touching controller code
- Different values per environment (Development vs Production)

---

### 7.3 `IFormFile` and multipart upload

HTML form sends:

```
Content-Type: multipart/form-data

file = (binary pptx data)
```

ASP.NET maps that to `IFormFile`. The field name **must** be `file` (matches frontend).

---

### 7.4 COM Interop on Windows

**What:** Calling PowerPoint as an external application.

**Rules:**
- Windows only
- PowerPoint must be installed
- Use STA thread
- Always call `Quit()` and release COM objects
- Expect Office dialogs/security restrictions on some machines

---

### 7.5 CORS

**What:** Browser security feature.

Frontend at `https://myapp.vercel.app` calling API at `http://localhost:5025` is a **cross-origin** request.

API must respond with `Access-Control-Allow-Origin` header — handled by `AddCors()` in `Program.cs`.

---

## 8. 🐛 Common problems and how to debug

| 🔴 Problem | 🔍 Likely cause | ✅ What to check |
|------------|-----------------|------------------|
| Upload works, no slides | PowerPoint not installed / IIS identity | [§10.7](#107-powerpoint--slide-export-on-the-server) |
| `Slide export failed unexpectedly` | Permissions or COM under IIS | App pool `Modify` on `Files\`; test API outside IIS |
| `Microsoft PowerPoint is not installed` | No desktop PowerPoint on server | Install Office PowerPoint (not web-only) |
| `Hiding the application window is not allowed` | Office policy blocks hidden automation | Do not set `Application.Visible = false` |
| `Presentation cannot be modified` | Opened pptx as read-only | Use read-write open flags |
| Port 5025 already in use | Old `dotnet run` still running | `netstat -ano \| findstr :5025` then kill PID |
| CORS error in browser | Frontend origin not in config | Add URL to `Cors:AllowedOrigins` — [§10.6](#106-cors--browser-errors-we-fixed) |
| `more-private address space local` | Public site calling `192.168.x.x` API | Use same network or public API URL — [§10.6](#106-cors--browser-errors-we-fixed) |
| HTTP **500.19** on IIS | Missing Hosting Bundle / bad `web.config` | [§10.2](#102-fix-http-error-50019) |
| **404** on site root `/` | No route at `/` (normal for Web API) | Test `/api/presentation/view` instead — [§10.4](#104-404-on-root-url-is-normal) |
| `Failed to determine the https port` | HTTPS redirect in HTTP-only dev | Keep HTTPS redirect disabled in Development |
| PowerPoint stuck in Task Manager | COM not cleaned up | Check `finally` block runs; kill `POWERPNT.EXE` |

### 🔎 Where to look first when debugging

| Priority | Where | What to look for |
|----------|-------|------------------|
| 1️⃣ | Terminal / IIS logs | `ILogger` output from controller and service |
| 2️⃣ | Upload response JSON | `slidesRendered` and `renderWarning` |
| 3️⃣ | `Files/slides/` folder | Are PNG files created after upload? |
| 4️⃣ | Task Manager | Is `POWERPNT.EXE` left running? |
| 5️⃣ | Browser DevTools → Network | CORS headers, redirects, status codes |

---

## 9. ✅ Junior developer TODO checklist

Use this checklist to learn the codebase hands-on. Do each task without AI assistance.

### Setup tasks

- [ ] Install .NET 10 SDK
- [ ] Clone/open the project
- [ ] Run `dotnet run` inside `backend/PresentationApi`
- [ ] Confirm API responds at `http://localhost:5025/api/presentation/view`

### Understanding tasks

- [ ] Draw a diagram of the upload flow on paper
- [ ] Explain to a teammate what DI means using `ISlideRenderService` as an example
- [ ] Find where the 50 MB upload limit is enforced (there are 2 places — can you find both?)
- [ ] Explain why we use an STA thread for PowerPoint

### Coding tasks (small)

- [ ] Add a log message when each slide PNG is exported successfully
- [ ] Add a new field `UploadedBy` to `UploadResponse` (hardcode `"system"` first)
- [ ] Change slide export resolution to 1280×720 in `appsettings.json` and verify output size changes
- [ ] Add validation: reject upload if filename contains spaces

### Coding tasks (medium)

- [ ] Add `GET /api/presentation/health` endpoint that returns `{ "status": "ok", "powerPointInstalled": true/false }`
- [ ] Add a config value `MaxSlideCount` and reject uploads that produce too many slides
- [ ] Store upload timestamp in a small `metadata.json` file alongside `current.pptx`

### Coding tasks (advanced)

- [ ] Create a second implementation of `ISlideRenderService` using LibreOffice CLI (if installed) as fallback
- [ ] Add unit tests for file validation logic (extract validation to a separate class first)
- [ ] Add authentication to the upload endpoint (API key in header)

### Verification tasks

- [ ] Upload a valid `.pptx` and confirm PNG files appear in `Files/slides/`
- [ ] Upload a `.txt` file renamed to `.pptx` and confirm rejection (or handle gracefully)
- [ ] Upload a file larger than 50 MB and confirm `400 Bad Request`
- [ ] Call `GET /api/presentation/slides/1` in browser and confirm image displays

### Deployment tasks

- [ ] Publish API with `dotnet publish` (not `bin\Release` copy)
- [ ] Install .NET 10 Hosting Bundle on server
- [ ] Configure IIS app pool → **No Managed Code**
- [ ] Add frontend origin to `Cors:AllowedOrigins` in server `appsettings.json`
- [ ] Set `Presentation:PublicBaseUrl` on server
- [ ] Grant app pool **Modify** permission on `Files\`
- [ ] Verify `GET /api/presentation/view` returns JSON (not 404 on `/`)

---

## 10. Publish to IIS (Windows server)

> 🚀 **When to use this section:** You are deploying the API to a Windows server with **IIS** (not `dotnet run` for production).  
> This documents real issues we hit during deployment and how we fixed them.

### 10.1 📋 Prerequisites checklist

| Requirement | Status | Notes |
|-------------|--------|-------|
| Windows Server / Windows 10+ | ✅ Required | PowerPoint COM only works on Windows |
| IIS installed | ✅ Required | Install **before** Hosting Bundle |
| **.NET 10 Hosting Bundle** | ✅ Required | [Download](https://dotnet.microsoft.com/download/dotnet/10.0) → **Hosting Bundle** |
| **Microsoft PowerPoint** (desktop) | ✅ Required | For slide PNG export — web-only Office does **not** work |
| Firewall port open | ✅ Required | e.g. port `85` for API, port `60` for frontend |

```
┌─────────────────────────────────────────────────────────────┐
│  🖥️  Windows Server                                         │
│  ┌─────────────────┐    ┌─────────────────────────────┐    │
│  │  IIS :60        │    │  IIS :85                    │    │
│  │  Frontend       │───▶│  PresentationApi (.NET 10)  │    │
│  │  (static HTML)  │    │  + PowerPoint COM export    │    │
│  └─────────────────┘    └─────────────────────────────┘    │
│         ▲                            │                       │
│         │                            ▼                       │
│    Browser on LAN              Files/current.pptx            │
│    192.168.15.100:60           Files/slides/*.png            │
└─────────────────────────────────────────────────────────────┘
```

---

### 10.2 🔧 Publish the API (correct way)

> ⚠️ **Do not** copy `bin\Release\net10.0` manually. Always use `dotnet publish` — it generates `web.config` and all files IIS needs.

```powershell
cd backend\PresentationApi
dotnet publish -c Release -o C:\inetpub\YunuscoSolarTV\api
```

**After publish, verify these files exist:**

| File | Why it matters |
|------|----------------|
| ✅ `web.config` | Tells IIS to use ASP.NET Core Module |
| ✅ `PresentationApi.dll` | Main application |
| ✅ `appsettings.json` | CORS, paths, `PublicBaseUrl` |
| ✅ `PresentationApi.exe` | App host (optional but normal) |

---

### 10.2.1 🛠️ Fix HTTP Error 500.19

**Symptom:**

```
HTTP Error 500.19 - Internal Server Error
The requested page cannot be accessed because the related configuration data for the page is invalid.
```

**Root causes we encountered:**

| Cause | Error code (typical) | Fix |
|-------|----------------------|-----|
| **ASP.NET Core Hosting Bundle not installed** | `0x8007000d` | Install [.NET 10 Hosting Bundle](https://dotnet.microsoft.com/download/dotnet/10.0), then `iisreset` |
| **IIS installed after Hosting Bundle** | `0x8007000d` | Re-run installer → **Repair** → `iisreset` |
| **Missing `web.config`** | Config error on `web.config` | Re-publish with `dotnet publish` |
| **Wrong app pool** | Various | Set .NET CLR to **No Managed Code** |

**Install Hosting Bundle (server-side):**

1. Download **Hosting Bundle** from the .NET 10 download page (not Runtime alone)
2. Run installer **as Administrator**
3. Restart IIS:

```powershell
iisreset
```

**Verify installation:**

```powershell
dotnet --list-runtimes
# Expect: Microsoft.AspNetCore.App 10.0.x

Get-WebGlobalModule | Where-Object { $_.Name -like '*AspNetCore*' }
# Expect: AspNetCoreModuleV2
```

**IIS site settings:**

| Setting | Value |
|---------|-------|
| Physical path | Folder containing `web.config` (e.g. `C:\inetpub\YunuscoSolarTV\api`) |
| App pool → .NET CLR version | **No Managed Code** |
| App pool → Managed pipeline | Integrated |

---

### 10.3 🌐 Deploy the frontend

Copy the `frontend/` folder to an IIS site (e.g. port `60`).

Update `frontend/js/config.js`:

```javascript
apiBaseUrl: "http://192.168.15.100:85",  // API URL — same network as browser
```

> 💡 **Rule:** The URL the browser uses to reach the API must match how users access the server (all LAN IPs, or all public IPs). Do not mix public frontend + private API IP.

---

### 10.4 ✅ 404 on root URL is normal

| URL | Expected result |
|-----|-----------------|
| `http://192.168.15.100:85/` | ❌ **404** — no homepage (this is a Web API only) |
| `http://192.168.15.100:85/api/presentation/view` | ✅ **200** JSON — `{"exists":false}` or slide metadata |

A **404 on `/` means IIS and the Hosting Bundle are working** — there is simply no route at the root.

---

### 10.5 ⚙️ Server `appsettings.json`

Edit on the **server** (in the publish folder), then restart IIS / recycle app pool.

```json
{
  "Cors": {
    "AllowedOrigins": [
      "http://localhost:3000",
      "http://127.0.0.1:5500",
      "https://*.vercel.app",
      "http://192.168.15.100:60"
    ]
  },
  "Presentation": {
    "StoragePath": "Files/current.pptx",
    "SlidesPath": "Files/slides",
    "PublicBaseUrl": "http://192.168.15.100:85"
  }
}
```

| Setting | Purpose |
|---------|---------|
| `Cors:AllowedOrigins` | Browser allows cross-origin calls from frontend port |
| `Presentation:PublicBaseUrl` | Slide image URLs in `/view` JSON point to correct host |

---

### 10.6 🚫 CORS & browser errors we fixed

#### Error A — Private Network Access (public → LAN)

```
Access to fetch ... has been blocked by CORS policy:
The request client is not a secure context and the resource is in more-private address space `local`.
```

| | Address |
|---|---------|
| Frontend | `http://202.74.243.118:5593` (public) |
| API | `http://192.168.15.100:85` (private LAN) |

**Cause:** Chrome blocks public websites from calling private IPs (`192.168.x.x`).

**Fix (Option A — public internet):** Use public API URL in `config.js`:

```javascript
apiBaseUrl: "http://202.74.243.118:85"
```

**Fix (Option B — LAN / kiosk — what we used):** Both frontend and API on same LAN:

```
Frontend: http://192.168.15.100:60
API:      http://192.168.15.100:85
```

---

#### Error B — Missing CORS origin (LAN deployment)

```
No 'Access-Control-Allow-Origin' header is present on the requested resource.
```

**Cause:** Frontend origin `http://192.168.15.100:60` was not in `Cors:AllowedOrigins`.

**Fix:** Add the frontend URL to server `appsettings.json` → restart IIS.

**Verify in DevTools → Network → Response headers:**

```
Access-Control-Allow-Origin: http://192.168.15.100:60
```

---

### 10.7 🖼️ PowerPoint & slide export on the server

Upload can succeed while slide export fails. The API returns:

```
Presentation uploaded successfully. · Warning: Slide export failed unexpectedly.
```

| Message in `renderWarning` | Meaning |
|----------------------------|---------|
| `Microsoft PowerPoint is not installed on this machine.` | Install desktop PowerPoint |
| `PowerPoint could not export slides...` | COM / license error |
| `Slide export failed unexpectedly.` | Often **IIS permissions** or **COM under app pool identity** |

**Requirements:**

| Item | Detail |
|------|--------|
| PowerPoint | Desktop app installed and licensed on the server |
| Folder permissions | App pool identity needs **Modify** on `Files\` |

```powershell
icacls "C:\inetpub\YunuscoSolarTV\api\Files" /grant "IIS AppPool\YourAppPoolName:(OI)(CI)M"
```

**Test outside IIS** (run as logged-in user):

```powershell
cd C:\inetpub\YunuscoSolarTV\api
$env:ASPNETCORE_ENVIRONMENT = "Development"
.\PresentationApi.exe
```

| Result | Diagnosis |
|--------|-----------|
| Works in console, fails in IIS | IIS identity cannot run PowerPoint COM |
| Fails everywhere | PowerPoint not installed or file permission issue |

> ⚠️ Microsoft does not officially support Office automation under IIS service accounts. For production kiosks, consider running the API as a Windows Service under a user account that can run PowerPoint.

---

### 10.8 📝 Deployment quick checklist

```
[ ] IIS installed
[ ] .NET 10 Hosting Bundle installed → iisreset
[ ] dotnet publish -c Release -o <site folder>
[ ] web.config present in publish folder
[ ] App pool = No Managed Code
[ ] Cors:AllowedOrigins includes frontend URL (on server appsettings.json)
[ ] Presentation:PublicBaseUrl set (on server appsettings.json)
[ ] config.js apiBaseUrl points to API (same network as users)
[ ] App pool has Modify on Files\
[ ] PowerPoint installed on server
[ ] GET /api/presentation/view returns JSON
[ ] Upload from frontend works without CORS error
[ ] PNG files appear in Files/slides/ after upload
```

---

## 📡 Quick reference — endpoints

| Method | URL | Returns |
|--------|-----|---------|
| `POST` | `/api/presentation/upload` | `UploadResponse` JSON |
| `GET` | `/api/presentation/view` | `PresentationViewResponse` JSON |
| `GET` | `/api/presentation/file` | `.pptx` binary |
| `GET` | `/api/presentation/slides/{index}` | `image/png` |

---

## 🎓 Summary for juniors

| ❓ Question | 💬 Answer |
|-------------|-----------|
| What framework? | ASP.NET Core 10 Web API |
| How many NuGet packages? | One (`Microsoft.AspNetCore.OpenApi`) |
| Where is business logic? | `Services/PowerPointSlideRenderService.cs` |
| Where are HTTP routes? | `Controllers/PresentationController.cs` |
| Where is config? | `appsettings.json` → `Options/PresentationOptions.cs` |
| Why PowerPoint COM? | Correct fonts/layout; browsers can't render pptx properly |
| Hardest part? | COM + STA thread + proper cleanup |
| How to deploy? | [§10 Publish to IIS](#10-publish-to-iis-windows-server) |
| Start reading from? | `Program.cs` → `PresentationController.cs` → `PowerPointSlideRenderService.cs` |

---

*📅 Last updated: June 2026 — matches backend in `backend/PresentationApi/`*
