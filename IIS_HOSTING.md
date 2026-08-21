# Hosting the Ocean Logistics Control Tower on IIS

## Short answer

**Yes — it will run on IIS, and Excel upload will work**, if you host this as a **Python/Flask app**, not as a static website.

| Host | Dashboard UI | Excel upload / refresh | Real shipment data |
|------|----------------|------------------------|--------------------|
| GitHub Pages | Yes (UI only) | **No** | No |
| `python run.py` on your PC | Yes | **Yes** | Yes |
| IIS serving this Flask app | Yes | **Yes** | Yes |
| IIS serving only the `static` folder | Yes (UI only) | **No** | No |

Upload works on IIS because Admin posts files to the Python backend. That backend writes into `data/`, runs the bundled ETL (`etl_engine/main.py`), saves `data/output/Consolidated_Ocean_DSR.xlsx`, and publishes `dashboard-data.json`. GitHub Pages cannot do any of that.

---

## Fast path (recommended)

Scripts under `deploy\iis\` copy the app to `C:\inetpub\ocean\logistics_website`, start Waitress on `127.0.0.1:8050`, and create an IIS site that reverse-proxies to it.

### 1. One-time downloads (if missing)

Install these on the Windows machine (reboot if prompted):

- [URL Rewrite](https://www.iis.net/downloads/microsoft/url-rewrite)
- [Application Request Routing (ARR)](https://www.iis.net/downloads/microsoft/application-request-routing)

Python 3.11+ should already be on PATH.

### 2. Run the installer as Administrator

Open **PowerShell as Administrator**:

```powershell
cd "C:\Users\cshah\OneDrive - Virginia Transformer Corp\Project\ocean\logistics_website"
powershell -ExecutionPolicy Bypass -File .\deploy\iis\Install-OceanIIS.ps1
```

Optional port change:

```powershell
.\deploy\iis\Install-OceanIIS.ps1 -SitePort 80
```

### 3. Open the site

- Dashboard: `http://localhost:8080/`
- Admin upload: `http://localhost:8080/admin`

Daily Excel updates: upload in **Admin** on that IIS URL. You do **not** republish to GitHub. The URL stays the same.

### 4. After code changes

Re-run the installer (it preserves `data\`), or copy files into `C:\inetpub\ocean\logistics_website` and restart the scheduled task:

```powershell
Start-ScheduledTask -TaskName OceanLogisticsWaitress
```

---

## What must be true for IIS upload to work

1. IIS is reverse-proxying (or running) the **Flask app**, not only HTML/CSS/JS.
2. The server has **Python 3** plus the packages in `requirements.txt` (including **Waitress**).
3. Copy **this one folder** (`logistics_website`) onto the server (prefer `C:\inetpub\...`, not OneDrive). ETL lives at `etl_engine/main.py`.
4. The account that runs Python can **read and write** `logistics_website\data\...`.
5. Request size is large enough (~100 MB).
6. Request timeout is long enough (ETL can take 1–2 minutes).
7. Users open the **IIS** URL `/admin`, not the GitHub Pages URL.

---

## Recommended layout on the server

```
C:\inetpub\ocean\logistics_website\     ← app + ETL + data
C:\inetpub\ocean\iis_site\web.config    ← IIS reverse-proxy only
```

Do **not** point IIS at `static\` only.

---

## Manual steps (if you prefer not to use the script)

### 1. Prerequisites

- IIS with Default Document, Static Content, **URL Rewrite**, **ARR**
- Python 3.11+ on PATH

### 2. Copy and install

```powershell
robocopy "<your-repo>\logistics_website" "C:\inetpub\ocean\logistics_website" /MIR /XD .git .venv __pycache__
cd C:\inetpub\ocean\logistics_website
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
```

### 3. Confirm Waitress without IIS

```powershell
.\deploy\iis\Start-OceanWaitress.ps1 -InstallRoot "C:\inetpub\ocean\logistics_website"
```

Open `http://127.0.0.1:8050` and `/admin`. Test an upload.

### 4. IIS reverse-proxy

1. Create a site whose physical path contains `deploy\iis\web.config` (or the copy under `C:\inetpub\ocean\iis_site`).
2. Enable ARR **Server Proxy Settings → Enable proxy**.
3. Binding example: `http://localhost:8080`.

IIS is the public face; Waitress does the work.

### 5. Keep Waitress running after logoff

Use the scheduled task created by `Install-OceanIIS.ps1` (`OceanLogisticsWaitress`), or an equivalent Windows service.

---

## Will upload work from my machine against the IIS site?

**Yes**, if you use the **IIS-hosted** URL:

- Browser → `http://your-server:8080/admin` → IIS → Waitress → ETL → `data/current` → dashboard refreshes.
- No GitHub republish after upload.
- Other users see new data after a page refresh.

**No**, if you use GitHub Pages, or IIS pointed only at `static\`.

---

## GitHub vs IIS

| Place | Purpose |
|-------|---------|
| GitHub | Source backup / code |
| GitHub Pages | Optional UI demo (no live data, no upload) |
| IIS (this app) | Real company site: dashboard + Admin + Excel |

Daily operations should use the **IIS URL**.

---

## Checklist if something fails

- **502 Bad Gateway**: Waitress not running, or ARR proxy not enabled. Start task `OceanLogisticsWaitress`; enable ARR proxy.
- Dashboard loads but Admin errors with HTML/JSON parse: IIS is not forwarding `/admin` and `/api` to Python.
- Upload times out: raise IIS and Waitress timeouts (script sets ~5 minutes).
- Upload rejected as too large: `maxAllowedContentLength` (script sets 100 MB).
- ETL import errors: confirm `etl_engine\main.py` is inside the install folder.
- Permission errors: `data\` needs Modify for the account running Waitress / IIS.
- Charts empty after upload: confirm `data\current\dashboard-data.json` updated, then hard-refresh.
- Excel file location: `data\output\Consolidated_Ocean_DSR.xlsx`.

---

## What you do not need

- You do not need GitHub Pages for IIS.
- You do not need to rebuild or redeploy after each Excel upload.
- You should not host this as a static-only IIS site if you want Admin to work.
