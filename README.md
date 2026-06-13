# YunuscoSolarTV — PowerPoint Upload & Slideshow Viewer

A simple web app for uploading a PowerPoint (`.pptx`) file and displaying it as an auto-playing slideshow — designed for TV/kiosk-style presentation displays (e.g. solar dashboard screens).

The project has two parts:

| Part | Stack | Role |
|------|-------|------|
| **Backend** | .NET 10 Web API | Accepts uploads, stores the file, exports slides to PNG via PowerPoint |
| **Frontend** | HTML / CSS / JavaScript | Upload page + fullscreen slideshow viewer (deployable to Vercel) |

---

## How it works

```
┌─────────────┐     upload .pptx      ┌──────────────────┐
│   Browser   │ ───────────────────►  │  .NET Web API    │
│  (Frontend) │                       │                  │
│             │ ◄── slide metadata ── │  Saves pptx      │
│             │                       │  Exports PNGs    │
│             │ ◄── slide images ──── │  via PowerPoint  │
└─────────────┘                       └──────────────────┘
```

1. **Upload** — User selects a `.pptx` file on the upload page. The API saves it as `Files/current.pptx` (replacing any previous file).
2. **Export** — On upload, the API uses **Microsoft PowerPoint COM automation** to export each slide as a 1920×1080 PNG into `Files/slides/`.
3. **View** — The view page fetches slide metadata and displays the PNGs as a slideshow with **5-second auto-advance** and **fullscreen** support.

Slides are rendered server-side by PowerPoint itself, so fonts and layout match what you see in the PowerPoint desktop app.

---

## Project structure

```
YunuscoSolarTV/
├── backend/
│   └── PresentationApi/          # .NET Web API
│       ├── Controllers/
│       │   └── PresentationController.cs
│       ├── Services/
│       │   ├── PowerPointSlideRenderService.cs   # COM export
│       │   └── ISlideRenderService.cs
│       ├── Models/               # Request/response DTOs
│       ├── Options/                # appsettings binding
│       ├── Files/
│       │   ├── current.pptx        # (created on upload)
│       │   └── slides/             # slide-001.png, slide-002.png, ...
│       ├── Program.cs
│       └── appsettings.json
│
└── frontend/
    ├── index.html                  # Upload page
    ├── view.html                   # Slideshow viewer
    ├── css/style.css
    ├── js/
    │   ├── config.js               # API URL & viewer settings
    │   ├── api.js                  # API client helpers
    │   ├── upload.js
    │   └── view.js
    └── vercel.json                 # Vercel deployment config
```

---

## Prerequisites

### Backend
- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- **Windows** (slide export uses PowerPoint COM)
- **Microsoft PowerPoint** installed and licensed on the machine running the API

### Frontend
- Any static file server for local dev (e.g. VS Code Live Server, `npx serve`)
- Optional: [Vercel](https://vercel.com) account for deployment

---

## Quick start (local)

### 1. Start the API

```powershell
cd backend\PresentationApi
dotnet run
```

The API listens on **http://localhost:5025** by default.

### 2. Serve the frontend

Open the `frontend/` folder with a static server. For example:

```powershell
cd frontend
npx serve .
```

Or use VS Code **Live Server** and open `index.html`.

### 3. Configure the frontend

Edit `frontend/js/config.js`:

```javascript
window.APP_CONFIG = {
  apiBaseUrl: "http://localhost:5025",  // your API base URL
  autoplayIntervalMs: 5000,             // slide interval (ms)
  autoFullscreen: true                    // attempt fullscreen on view page load
};
```

### 4. Upload and view

1. Open **Upload** page → choose a `.pptx` → click **Upload**
2. Wait for the success message (should include slide count, e.g. `5 slides rendered`)
3. Open **View** page → slideshow auto-plays every 5 seconds

---

## API reference

Base URL: `http://localhost:5025` (local) or your deployed API URL.

### `POST /api/presentation/upload`

Upload a `.pptx` file. Replaces any existing presentation.

| Item | Value |
|------|-------|
| Content-Type | `multipart/form-data` |
| Field name | `file` |
| Max size | 50 MB |
| Allowed type | `.pptx` only |

**Success (200)**

```json
{
  "message": "Presentation uploaded successfully.",
  "fileName": "current.pptx",
  "fileSizeBytes": 1048576,
  "uploadedAtUtc": "2026-06-13T10:00:00Z",
  "slideCount": 12,
  "slidesRendered": true,
  "renderWarning": null
}
```

**Error (400)**

```json
{
  "error": "Invalid file type.",
  "details": "Only .pptx files are allowed."
}
```

---

### `GET /api/presentation/view`

Returns metadata about the current presentation and slide URLs.

**No file uploaded (200)**

```json
{
  "exists": false,
  "fileName": null,
  "fileSizeBytes": null,
  "lastModifiedUtc": null,
  "fileUrl": null,
  "officeViewerUrl": null,
  "slidesAvailable": false,
  "slideCount": 0,
  "slides": []
}
```

**File exists (200)**

```json
{
  "exists": true,
  "fileName": "current.pptx",
  "fileSizeBytes": 1048576,
  "lastModifiedUtc": "2026-06-13T10:00:00Z",
  "fileUrl": null,
  "officeViewerUrl": null,
  "slidesAvailable": true,
  "slideCount": 12,
  "slides": [
    { "index": 1, "url": "http://localhost:5025/api/presentation/slides/1" },
    { "index": 2, "url": "http://localhost:5025/api/presentation/slides/2" }
  ]
}
```

---

### `GET /api/presentation/file`

Returns the raw `.pptx` binary.

| Response | Details |
|----------|---------|
| 200 | `Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation` |
| 404 | No presentation uploaded yet |

---

### `GET /api/presentation/slides/{index}`

Returns a single slide image (PNG).

| Parameter | Description |
|-----------|-------------|
| `index` | 1-based slide number |

| Response | Details |
|----------|---------|
| 200 | `Content-Type: image/png` |
| 404 | Slide or presentation not found |

---

## Configuration

### Backend — `appsettings.json`

```json
{
  "Cors": {
    "AllowedOrigins": [
      "http://localhost:3000",
      "http://127.0.0.1:5500",
      "https://*.vercel.app"
    ]
  },
  "Presentation": {
    "StoragePath": "Files/current.pptx",
    "SlidesPath": "Files/slides",
    "SlideExportWidth": 1920,
    "SlideExportHeight": 1080,
    "MaxFileSizeBytes": 52428800,
    "AllowedExtensions": [ ".pptx" ],
    "PublicBaseUrl": ""
  }
}
```

| Setting | Purpose |
|---------|---------|
| `StoragePath` | Where the uploaded `.pptx` is saved |
| `SlidesPath` | Folder for exported PNG slides |
| `SlideExportWidth/Height` | Resolution of exported slide images |
| `MaxFileSizeBytes` | Upload size limit (default 50 MB) |
| `PublicBaseUrl` | Public HTTPS URL of the API (for production slide URLs) |
| `Cors:AllowedOrigins` | Frontend origins allowed to call the API |

### Frontend — `js/config.js`

| Setting | Default | Purpose |
|---------|---------|---------|
| `apiBaseUrl` | `http://localhost:5025` | Backend API base URL |
| `autoplayIntervalMs` | `5000` | Milliseconds between slides |
| `autoFullscreen` | `true` | Try to enter fullscreen when view page loads |

---

## Deployment

### Backend

Deploy `backend/PresentationApi` to a **Windows server** with PowerPoint installed (e.g. Azure App Service on Windows, IIS, or a local TV/kiosk PC).

1. Publish the API:
   ```powershell
   dotnet publish -c Release -o ./publish
   ```
2. Set `Presentation:PublicBaseUrl` to your public API URL (e.g. `https://api.yourdomain.com`)
3. Add your Vercel frontend domain to `Cors:AllowedOrigins`

### Frontend (Vercel)

1. Set the Vercel project root to the `frontend/` folder
2. Update `frontend/js/config.js` → `apiBaseUrl` to your deployed API URL
3. Deploy — `vercel.json` is already included

---

## Viewer features

| Feature | Description |
|---------|-------------|
| **Auto-play** | Advances to the next slide every 5 seconds (configurable) |
| **Pause / Play** | Toggle auto-advance |
| **Prev / Next** | Manual navigation (resets the auto-play timer) |
| **Fullscreen** | Shows slides only — hides header, toolbar, and metadata |
| **Loop** | After the last slide, returns to slide 1 |

---

## Troubleshooting

### Upload succeeds but no slides on View page

- **PowerPoint must be installed** on the machine running the API
- Re-upload the file after fixing PowerPoint — slides are generated at upload time
- Check the upload response for `renderWarning`

### `PowerPoint COM export failed`

Common causes on Windows:

| Error | Fix |
|-------|-----|
| `Hiding the application window is not allowed` | Fixed in current code — do not set `Application.Visible = false` |
| `Presentation cannot be modified` | File must be opened read-write (not read-only) for export |
| PowerPoint not installed | Install Microsoft PowerPoint on the API server |

Restart the API after code changes:

```powershell
# Ctrl+C to stop, then:
dotnet run
```

### Port already in use (`address already in use`)

Another instance is running on port 5025:

```powershell
netstat -ano | findstr :5025
Stop-Process -Id <PID> -Force
```

### CORS errors from frontend

Add your frontend origin to `Cors:AllowedOrigins` in `appsettings.json`.

### HTTPS redirect warning in development

HTTPS redirection is disabled in Development mode. This warning should not appear after restarting with the latest code.

---

## Design decisions (AI agent build notes)

This project was built with the following architectural choices:

1. **Server-side slide export via PowerPoint COM** — chosen over in-browser libraries (e.g. PptxViewJS) because browser renderers cannot reliably reproduce embedded/custom fonts. PowerPoint exports pixel-accurate PNGs.

2. **Single-file replace model** — only one `current.pptx` is kept at a time, simplifying TV/kiosk use cases where one presentation loops continuously.

3. **Vanilla HTML/CSS/JS frontend** — no build step required; easy to deploy on Vercel as static files.

4. **Image-based slideshow** — PNG slides enable simple auto-play, fullscreen, and consistent rendering across browsers and TV displays.

---

## Tech stack summary

| Layer | Technology |
|-------|------------|
| API framework | ASP.NET Core 10 Web API |
| Slide export | PowerPoint COM Interop (Windows) |
| File storage | Local filesystem (`Files/`) |
| Frontend | Vanilla HTML, CSS, JavaScript |
| Frontend hosting | Vercel (static) |
| CORS | ASP.NET Core CORS middleware |
