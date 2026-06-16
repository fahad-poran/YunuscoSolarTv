# ☀️ Yunusco Solar TV

PowerPoint upload and slideshow viewer for TV/kiosk displays.

This app lets an admin upload a `.pptx` file, converts each slide into PNG images on the server, and shows the slides in a fullscreen browser viewer together with the SolScada dashboard.

---

## 📌 What This App Does

| Area | Description |
| --- | --- |
| 🖥️ Frontend | Static HTML, CSS, and JavaScript upload/view pages |
| ⚙️ Backend | ASP.NET Core Web API |
| 📊 PowerPoint Rendering | Uses installed Microsoft PowerPoint on Windows to export slides as PNG |
| 📺 TV Viewer | Shows SolScada dashboard first, then uploaded slides in a loop |
| 🌐 Hosting | Frontend can run from IIS/static hosting; backend runs on Windows/IIS |

---

## 🧱 Project Structure

```text
YunuscoSolarTV/
├── backend/
│   └── PresentationApi/
│       ├── Controllers/
│       │   └── PresentationController.cs
│       ├── Services/
│       │   ├── PowerPointSlideRenderService.cs
│       │   └── ISlideRenderService.cs
│       ├── Models/
│       ├── Options/
│       ├── Files/
│       │   ├── current.pptx
│       │   └── slides/
│       ├── Program.cs
│       ├── appsettings.json
│       └── appsettings.Production.json
│
└── frontend/
    ├── index.html
    ├── view.html
    ├── css/
    │   └── style.css
    └── js/
        ├── config.js
        ├── api.js
        ├── upload.js
        └── view.js
```

---

## 🔄 How The App Works

```text
Browser Upload Page
        │
        │ uploads .pptx
        ▼
ASP.NET Core API
        │
        │ saves file as Files/current.pptx
        ▼
PowerPoint COM Automation
        │
        │ exports each slide as PNG
        ▼
Files/slides/slide-001.png, slide-002.png, ...
        │
        ▼
Browser View Page
```

The backend uses real Microsoft PowerPoint to render slides. This is important because PowerPoint gives better output for custom fonts, layouts, charts, and embedded objects than most browser-only PPTX viewers.

---

## ✅ Server Requirements

Before publishing, the Windows server should have:

| Requirement | Why It Is Needed |
| --- | --- |
| ✅ IIS | Hosts the backend API and/or frontend |
| ✅ ASP.NET Core Hosting Bundle | Allows IIS to run ASP.NET Core apps |
| ✅ Matching .NET runtime | Must match the project target framework, for example `.NET 10` for `net10.0` |
| ✅ Microsoft PowerPoint | Required to convert PPTX slides into PNG |
| ✅ Open firewall port | Allows other devices to call the API |
| ✅ Dedicated Windows user for PowerPoint | Avoids IIS App Pool COM permission issues |

---

## 🧪 Health Check

The backend has a health endpoint:

```text
GET /api/health
```

Example:

```text
http://192.168.15.6:75/api/health
```

Expected response:

```json
{
  "status": "running",
  "application": "PresentationApi",
  "environment": "Production",
  "serverTimeUtc": "2026-06-15T00:00:00+00:00"
}
```

If this URL does not work, fix the backend/server first. CORS cannot work until the API is reachable.

---

## 🚀 Easy Publish Guide For Junior Developers

Follow these steps in order.

### 1. Install The Correct .NET Hosting Bundle

Check the project target framework:

```xml
<TargetFramework>net10.0</TargetFramework>
```

If the project uses `net10.0`, install the **.NET 10 ASP.NET Core Hosting Bundle** on the server.

After installing, restart IIS:

```powershell
iisreset
```

Check installed runtimes:

```powershell
dotnet --list-runtimes
```

You should see:

```text
Microsoft.NETCore.App 10.x.x
Microsoft.AspNetCore.App 10.x.x
```

> 💡 If the server shows `HTTP Error 500.31`, it usually means the required .NET runtime/hosting bundle is missing.

---

### 2. Publish The Backend

Open PowerShell as **Administrator**.

Go to the backend project:

```powershell
cd D:\YunuscoSolarTV\backend\PresentationApi
```

Stop IIS before publishing if files are locked:

```powershell
iisreset /stop
```

Publish:

```powershell
dotnet publish -c Release -o C:\inetpub\YunuscoSolarTV\api
```

Start IIS again:

```powershell
iisreset /start
```

> 💡 If publish fails with `Access to the path is denied`, run PowerShell as Administrator and stop the IIS site/app pool before publishing.

---

### 3. Configure IIS Site Binding

In **IIS Manager**:

1. Select the API site.
2. Click **Bindings...**
3. Add or edit:

```text
Type: http
IP address: All Unassigned
Port: 75
Host name: empty
```

Then browse on the server:

```text
http://localhost:75/api/health
```

Then browse from another computer:

```text
http://192.168.15.6:75/api/health
```

---

### 4. Allow Port 75 In Windows Firewall

Run PowerShell as **Administrator**:

```powershell
New-NetFirewallRule `
  -DisplayName "YunuscoSolarTV API Port 75" `
  -Direction Inbound `
  -Protocol TCP `
  -LocalPort 75 `
  -Action Allow
```

Use this check:

| Test | Meaning |
| --- | --- |
| `http://localhost:75/api/health` works | IIS/API is running locally |
| `http://server-ip:75/api/health` fails | Firewall, IP binding, or network issue |
| Both fail | IIS/API/runtime issue |

---

### 5. Configure CORS

Backend CORS is configured in:

```text
backend/PresentationApi/appsettings.json
backend/PresentationApi/appsettings.Production.json
```

Example:

```json
{
  "Cors": {
    "AllowedOrigins": [
      "http://192.168.15.6:70",
      "http://localhost:*",
      "http://127.0.0.1:*"
    ]
  }
}
```

If the frontend runs on:

```text
http://192.168.15.6:70
```

then that exact origin must be listed in `Cors:AllowedOrigins`.

> 💡 CORS compares origin only: protocol + IP/domain + port.  
> Example: `http://192.168.15.6:70` and `http://192.168.15.6:75` are different origins.

---

### 6. Configure Frontend API URL

Edit:

```text
frontend/js/config.js
```

Set the backend URL:

```javascript
const apiBaseUrl = isLocalhost ? "http://localhost:5025" : "http://192.168.15.6:75";
```

If the server IP changes, update this file.

---

## 🔐 PowerPoint COM Permission Fix

This was the main upload/render issue after publishing.

The upload worked, but slide rendering failed with:

```text
PowerPoint COM access denied.
The IIS app pool identity cannot start PowerPoint.
```

### Why This Happens

PowerPoint opens fine when you are logged in, but IIS does not run as your desktop user by default.

IIS usually runs as:

```text
IIS AppPool\YourAppPoolName
```

That account cannot easily open desktop Office apps like PowerPoint.

### Recommended Fix

Run the IIS App Pool under a normal Windows user that has opened PowerPoint once.

---

### Step-By-Step PowerPoint User Setup

#### 1. Create A Windows User

Create a user, for example:

```text
pptuser
```

In **Computer Management**:

```text
Local Users and Groups > Users > New User
```

Recommended options:

```text
User must change password at next logon: unchecked
Password never expires: checked
```

---

#### 2. Log In As That User

Sign in to Windows Server as:

```text
.\pptuser
```

or:

```text
SERVERNAME\pptuser
```

---

#### 3. Open PowerPoint Once

Open Microsoft PowerPoint manually.

Accept/finish any first-run screens:

```text
License agreement
Activation/sign-in
Privacy settings
Welcome screen
Default file prompts
Protected View / Read Only prompts
```

Open a `.pptx` file once and make sure PowerPoint works normally.

---

#### 4. Disable Protected View For This User

In PowerPoint:

```text
File > Options > Trust Center > Trust Center Settings > Protected View
```

Disable Protected View options that block local/network files.

This prevents automation from getting stuck on a yellow warning bar or read-only prompt.

---

#### 5. Set IIS App Pool Identity

In **IIS Manager**:

1. Go to **Application Pools**.
2. Select the API app pool.
3. Click **Advanced Settings**.
4. Find **Identity**.
5. Select **Custom account**.
6. Enter:

```text
.\pptuser
```

7. Enter the password.
8. Recycle the app pool.

---

#### 6. Give Folder Permission

Give `pptuser` permission to the API folder:

```text
C:\inetpub\YunuscoSolarTV\api
```

Required permissions:

```text
Read
Write
Modify
```

PowerShell example:

```powershell
icacls "C:\inetpub\YunuscoSolarTV\api" /grant "pptuser:(OI)(CI)M"
```

---

#### 7. Grant DCOM Permission If Needed

If PowerPoint still says COM access denied:

1. Press `Win + R`.
2. Run:

```text
dcomcnfg
```

3. Go to:

```text
Component Services
  Computers
    My Computer
      DCOM Config
```

4. Find:

```text
Microsoft PowerPoint Presentation
```

or:

```text
Microsoft PowerPoint Application
```

5. Right-click > **Properties**.
6. Go to **Security**.
7. Under **Launch and Activation Permissions**, choose **Customize** > **Edit**.
8. Add `pptuser`.
9. Allow:

```text
Local Launch
Local Activation
```

10. Under **Access Permissions**, allow:

```text
Local Access
```

11. Recycle the IIS app pool.

---

## 🧯 Common Problems And Fixes

| Problem | Cause | Fix |
| --- | --- | --- |
| `HTTP Error 500.31` | Missing .NET runtime/hosting bundle | Install matching ASP.NET Core Hosting Bundle |
| Publish says `Access denied` | IIS is locking files or no admin permission | Stop IIS/app pool and publish as Administrator |
| `/api/health` does not open | API not running, wrong port, firewall, or IIS binding | Check IIS binding, firewall, and runtime |
| CORS error in browser | Frontend origin is not allowed by backend | Add frontend origin to `Cors:AllowedOrigins` |
| Upload succeeds but slides fail | PowerPoint COM cannot start | Use `pptuser` as app pool identity and configure DCOM |
| PowerPoint opens read-only/protected view | Office security prompt blocks automation | Disable Protected View for `pptuser` |
| IIS says WAS/W3SVC stopped | IIS services are stopped | Start `WAS` and `W3SVC` services |

---

## 🛠️ Useful Commands

Check installed .NET runtimes:

```powershell
dotnet --list-runtimes
```

Publish backend:

```powershell
cd D:\YunuscoSolarTV\backend\PresentationApi
dotnet publish -c Release -o C:\inetpub\YunuscoSolarTV\api
```

Restart IIS:

```powershell
iisreset
```

Stop IIS:

```powershell
iisreset /stop
```

Start IIS:

```powershell
iisreset /start
```

Start IIS services:

```powershell
Start-Service WAS
Start-Service W3SVC
```

Check port usage:

```powershell
netstat -ano | findstr :75
```

Allow firewall port:

```powershell
New-NetFirewallRule `
  -DisplayName "YunuscoSolarTV API Port 75" `
  -Direction Inbound `
  -Protocol TCP `
  -LocalPort 75 `
  -Action Allow
```

---

## 🧭 Recommended Publish Checklist

Before saying deployment is done, verify each item:

```text
[ ] .NET Hosting Bundle installed
[ ] IIS site created
[ ] IIS binding uses port 75
[ ] Windows firewall allows port 75
[ ] Backend published successfully
[ ] /api/health works on localhost
[ ] /api/health works from another computer
[ ] Frontend config.js points to correct API URL
[ ] Backend CORS allows frontend origin
[ ] PowerPoint installed
[ ] pptuser created
[ ] PowerPoint opened once as pptuser
[ ] Protected View handled/disabled for pptuser
[ ] IIS App Pool runs as pptuser
[ ] pptuser has Modify permission on API folder
[ ] Upload renders slides successfully
```

---

## 🧑‍💻 Local Development

Run the backend:

```powershell
cd D:\YunuscoSolarTV\backend\PresentationApi
dotnet run
```

Default local API:

```text
http://localhost:5025
```

Open the frontend:

```text
frontend/index.html
frontend/view.html
```

For local development, `frontend/js/config.js` automatically uses:

```text
http://localhost:5025
```

when opened from `localhost` or `127.0.0.1`.

---

## 📡 API Endpoints

| Method | Endpoint | Purpose |
| --- | --- | --- |
| `GET` | `/api/health` | Confirms backend is running |
| `POST` | `/api/presentation/upload` | Uploads `.pptx` and exports slides |
| `GET` | `/api/presentation/view` | Returns current presentation and slide metadata |
| `GET` | `/api/presentation/file` | Downloads current `.pptx` |
| `GET` | `/api/presentation/slides/{index}` | Returns one slide PNG |

---

## ✅ Final Notes

PowerPoint automation on a server is sensitive to user identity. If upload works but slide rendering fails, the project code is usually fine. The problem is normally Windows permissions, IIS App Pool identity, Office first-run prompts, or DCOM launch permissions.

The safest production setup for this app is:

```text
IIS App Pool Identity = dedicated Windows user
PowerPoint opened once as that same user
API folder writable by that same user
DCOM permission granted if Windows blocks COM launch
```
