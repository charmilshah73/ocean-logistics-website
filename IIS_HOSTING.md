# Hosting the Ocean Logistics Control Tower on IIS

## Short answer

**Yes — it will run on IIS, and Excel upload will work**, if you host this as a **Python/Flask app**, not as a static website.

| Host | Dashboard UI | Excel upload / refresh | Real shipment data |
|------|----------------|------------------------|--------------------|
| GitHub Pages | Yes (UI only) | **No** | No |
| `python run.py` on your PC | Yes | **Yes** | Yes |
| IIS serving this Flask app | Yes | **Yes** | Yes |
| IIS serving only the `static` folder | Yes (UI only) | **No** | No |

Upload works on IIS because Admin posts files to the Python backend. That backend writes into `data/`, runs the existing ETL (`v10_pkg`), and publishes `dashboard-data.json`. GitHub Pages cannot do any of that.

---

## What must be true for IIS upload to work

1. IIS is reverse-proxying (or running) the **Flask app**, not only HTML/CSS/JS.
2. The server has **Python 3** plus the packages in `requirements.txt`.
3. The **`v10_pkg` folder stays next to this project** (same parent as `logistics_website`). Upload calls that ETL. If `v10_pkg` is missing, upload will fail.
4. The IIS app-pool account can **read and write** these folders:
   - `logistics_website\data\input`
   - `logistics_website\data\staging`
   - `logistics_website\data\archive`
   - `logistics_website\data\current`
5. Request size is large enough (this app allows about **80 MB** per file).
6. The request timeout is long enough. Consolidation can take **1–2 minutes**.
7. Users open the IIS site URL (intranet or localhost), then go to **`/admin`**, not the GitHub Pages URL.

If those are true, uploading from a browser against your IIS site **will run** the same pipeline as local `python run.py`.

---

## Recommended layout on the server

Keep this structure (do not flatten it):

```
C:\inetpub\ocean\
  logistics_website\     ← this repo (Flask app)
  v10_pkg\               ← existing ETL package (required for upload)
```

Do **not** put the site root on `static\` only. The site root must be `logistics_website`.

---

## Steps to host on IIS

### 1. Install prerequisites on the Windows server (or your PC)

- IIS with:
  - Default Document
  - Static Content
  - **URL Rewrite**
  - **Application Request Routing (ARR)** if you reverse-proxy to Python
- Python 3.11 or 3.12 (64-bit), and check “Add python.exe to PATH”
- Git (optional, to pull from GitHub)

### 2. Copy the app onto the machine

Put `logistics_website` and `v10_pkg` on a local disk (for example `C:\inetpub\ocean\`). Avoid relying on OneDrive sync for the live site — it can lock files during upload.

In `logistics_website`:

- Create a virtual environment
- Install `requirements.txt`
- Also install a production WSGI server such as **Waitress** (Flask’s built-in server is for development only)

### 3. Confirm it runs without IIS first

From `logistics_website`, start the app bound to all interfaces (host `0.0.0.0`, port `8050`).

Open:

- `http://127.0.0.1:8050` — dashboard
- `http://127.0.0.1:8050/admin` — upload

Do a test upload here. If this fails, IIS will also fail. Fix Python / `v10_pkg` / folder permissions first.

### 4. Create an IIS site that reverse-proxies to Python

This is the usual, reliable pattern:

1. Create a Windows service or scheduled task that **always runs Waitress** (or equivalent) on `127.0.0.1:8050` with working directory `logistics_website`.
2. In IIS, create a site (example bindings: `http://logistics.internal:80` or `http://localhost:8080`).
3. Install **URL Rewrite + ARR**.
4. Enable ARR proxy.
5. Add a reverse-proxy rule: all requests to this IIS site go to `http://127.0.0.1:8050`.

IIS then acts as the public face. Python still does the work. `/admin` upload hits the same Flask routes as local dev.

Alternative (more IIS-native): **HttpPlatformHandler** so IIS starts `python` / Waitress itself. Same idea — IIS must start Python, not only serve files.

### 5. Set IIS limits so large Excel files are accepted

In IIS for this site:

- Increase **maxAllowedContentLength** (suggest 100 MB or more)
- Increase **uploadReadAheadSize** if large posts fail early
- Increase **connection / request timeout** to at least 3–5 minutes (ETL is slow)

Also raise the Python/Waitress timeout if you set one.

### 6. Folder permissions

Give the identity that runs Python (and the IIS app-pool identity if it writes files) **Modify** on:

`logistics_website\data`

If upload returns a permission error, this is the usual cause.

### 7. Environment / binding

- Listen on `0.0.0.0` if other PCs on the network will use the site.
- Keep the process running after you log off (Windows service, not a leftover Command Prompt).
- Open the Windows Firewall port if colleagues will use the IIS hostname.

### 8. Smoke test after IIS is in front

From a browser:

1. Open the IIS site home page — landing + Ocean dashboard should load.
2. Open `/admin`.
3. Upload Liyana DSR + GoComet/Detailed Tracking (plus optional files).
4. Wait until it says the dashboard refreshed.
5. Go back to the dashboard and confirm KPIs/charts update **without** redeploying IIS.

If step 3–4 work, hosting is correct.

---

## Will upload work from my machine against the IIS site?

**Yes**, if you are talking to **your IIS-hosted Flask app**.

- Browser on your PC → `http://your-iis-server/admin` → IIS → Python → ETL → `data/current` → dashboard refreshes.
- You do **not** need to republish to GitHub after an upload.
- Other users hitting the same IIS URL will see the new data after a page refresh.

**No**, if you:

- Open the GitHub Pages URL and upload there
- Point IIS only at the `static` folder (no Python)
- Run IIS on a machine that cannot see `v10_pkg`
- Run IIS from a synced OneDrive copy that locks `data\` files

---

## GitHub vs IIS — use both this way

| Place | Purpose |
|-------|---------|
| GitHub | Source backup / code |
| GitHub Pages | Optional UI demo (no live data, no upload) |
| IIS (this app) | Real company site: dashboard + Admin upload |

Daily operations should use the **IIS URL**, not GitHub Pages.

---

## Checklist if something fails

- Dashboard loads but Admin errors with HTML/JSON parse: IIS is not forwarding `/admin` and `/api` to Python.
- Upload starts then times out: raise IIS and Python timeouts.
- Upload rejected as too large: raise IIS request size.
- “Missing v10_pkg” / import errors: copy `v10_pkg` next to `logistics_website`.
- Refresh succeeds locally but not on IIS: app-pool / Python service cannot write to `data\`.
- Charts empty after upload: confirm `data\current\dashboard-data.json` was updated, then hard-refresh the browser.

---

## What you do not need

- You do not need GitHub Pages for IIS.
- You do not need to rebuild or redeploy after each Excel upload.
- You should not host this as a static-only IIS site if you want Admin to work.
