# Granny

A lightweight, multi-user family history manager that runs entirely inside Google Workspace — no server, no database subscription, no build pipeline. Google Sheets is the database, Google Apps Script is the backend API, and a single HTML file delivers the Vue 3 frontend.

---

## Features

- **People profiles** — names, birth/death dates, places, notes, living flag
- **Complex relationships** — biological, adopted, step, and foster links; multiple marriages and blended families
- **Shared events** — a single Census, immigration, or military record can link to any number of people, each with their own role (subject, spouse, witness, informant…)
- **Interactive pedigree tree** — D3-rendered SVG canvas with pan, zoom, and click-to-navigate between generations
- **Multi-user access control** — admin / editor / viewer roles stored in a Users sheet; roles are enforced on every write via `LockService`-protected server functions
- **No build step** — Vue 3, Tailwind CSS, and D3.js all load from CDN; the entire frontend is one HTML file

---

## Tech Stack

| Layer | Technology |
|---|---|
| Database | Google Sheets (6 tabs) |
| Backend API | Google Apps Script (`Code.gs`) |
| Frontend | Vue 3 (Options API, CDN) |
| Styling | Tailwind CSS (Play CDN) |
| Visualisation | D3.js v7 (CDN) |
| Auth | Google OAuth via `Session.getActiveUser()` |

---

## Project Files

| File | Purpose |
|---|---|
| `Code.gs` | Apps Script backend — HTTP entry points, router, all CRUD functions |
| `Index.html` | Complete single-page frontend — Vue app, D3 tree, modals, styles |
| `deploy.py` | Automated setup script — creates the spreadsheet, Apps Script project, and uploads all code |
| `create_template.py` | Standalone script that generates `granny_template.xlsx` (offline, no Google auth needed) |
| `schema.md` | Full column-level schema reference with FK diagram and JSON assembly notes |

---

## Deployment

### Prerequisites

- A Google account
- Python 3.9+ with `pip`
- Access to [Google Drive](https://drive.google.com)

---

### Step 1 — One-time Google Cloud Setup

`deploy.py` uses two Google APIs (Drive and Apps Script). You need to create credentials once.

1. Go to [console.cloud.google.com](https://console.cloud.google.com) and create or select a project.
2. Enable both APIs:
   - **APIs & Services → Library** → search `Google Drive API` → **Enable**
   - **APIs & Services → Library** → search `Apps Script API` → **Enable**
3. Create OAuth credentials:
   - **APIs & Services → Credentials → Create credentials → OAuth client ID**
   - Application type: **Desktop app** → **Create**
   - Click **Download JSON**, rename the file to `credentials.json`, and place it in the same folder as `deploy.py`.
4. Enable the Apps Script API for your Google account *(one-time per account)*:
   - Go to [script.google.com/home/usersettings](https://script.google.com/home/usersettings)
   - Toggle **Google Apps Script API** to **On**

---

### Step 2 — Install Dependencies

```bash
pip install openpyxl google-api-python-client google-auth-oauthlib
```

---

### Step 3 — Run `deploy.py`

```bash
python deploy.py
```

The script will:
- Open a browser window asking you to sign in with your Google account and grant access.
- Create the **Granny** spreadsheet on your Drive with all 6 tabs and headers.
- Create a container-bound Apps Script project named **Granny**.
- Upload `Code.gs` and `Index.html` into the project automatically.

When it finishes it prints two URLs:

```
Spreadsheet : https://docs.google.com/spreadsheets/d/...
Script editor: https://script.google.com/d/.../edit
```

> **Credentials are cached** in `token.json` after the first run. Subsequent runs skip the browser login.

> **Using `create_template.py` instead?** If you only want the XLSX without any Google API calls, run `python create_template.py` to generate `granny_template.xlsx`, then upload and convert it manually.

---

### Step 4 — Register Yourself as an Admin

The backend enforces roles on every write. Before the first deploy, add yourself to the `Users` sheet manually so the app recognises you.

1. Go back to your Google Sheet.
2. Click the **`Users`** tab.
3. In row 2, enter the following values in the matching columns:

| Column | Value |
|---|---|
| `user_id` | `USR0001` |
| `email` | your Google account email (e.g. `you@gmail.com`) |
| `display_name` | Your name |
| `role` | `admin` |
| `created_at` | today's date |

> The `email` must match exactly what `Session.getActiveUser().getEmail()` returns for your account — this is always your primary Google account email.

---

### Step 5 — Deploy as a Web App

1. In the Apps Script editor, click **Deploy → New deployment**.
2. Click the **gear icon** next to "Select type" and choose **Web app**.
3. Fill in the deployment settings:

| Setting | Value |
|---|---|
| **Description** | `v1` (or any label) |
| **Execute as** | `Me` (your Google account) |
| **Who has access** | `Anyone with a Google account` |

> **Why "Execute as: Me"?** The script reads and writes your spreadsheet. Running as you means it always has access, regardless of which user is viewing the app.

> **Why "Anyone with a Google account"?** This ensures `Session.getActiveUser().getEmail()` returns the viewer's real email, which is how roles are enforced. If you set it to "Anyone" (no login required), the email will be empty and all write operations will be blocked.

4. Click **Deploy**.
5. If prompted, click **Authorize access** and grant the requested permissions (read/write Sheets, identify you as the current user).
6. Copy the **Web app URL** — this is the live URL for your app.

---

### Step 6 — Open the App

Paste the Web app URL into your browser. You should see the Granny interface load with an empty sidebar ready for your first person.

---

### Updating the App After Code Changes

Google Apps Script deployments are versioned. After editing `Code.gs` or `Index.html`:

1. Click **Deploy → Manage deployments**.
2. Click the **pencil (edit) icon** on your existing deployment.
3. Change the version dropdown to **"New version"**.
4. Click **Deploy**.

> You must create a new version — editing the code alone does not update the live deployment.

---

### Adding More Users

To grant access to collaborators:

1. Add a row to the `Users` sheet with their Google email and the appropriate role:
   - `admin` — full access including deletes
   - `editor` — can create and update; cannot delete
   - `viewer` — read-only (write operations return a permission error)
2. Share the Web app URL with them. They will be prompted to sign in with their Google account on first visit.

> You do not need to share the underlying spreadsheet with collaborators. Only the `Users` sheet row (and the Web app URL) is needed.

---

## Architecture Notes

### How the frontend communicates with the backend

Because the HTML is served by the same Apps Script deployment, the frontend uses `google.script.run` rather than `fetch()`. This avoids CORS entirely and means no authentication tokens need to be managed client-side.

```js
// All Vue methods go through this wrapper
function callServer(action, data) {
  return new Promise((resolve, reject) => {
    google.script.run
      .withSuccessHandler(resolve)
      .withFailureHandler(reject)
      .execute({ action, data });
  });
}
```

### Concurrency and data integrity

Every write operation acquires a `LockService.getScriptLock()` before touching the spreadsheet and releases it in a `finally` block. This serialises concurrent writes across all users, preventing row-index drift during simultaneous edits.

### Pedigree payload

`getPedigreePayload(personId)` loads all six sheets into memory in a single batch, builds an in-memory index of people by ID, then resolves all foreign keys without further sheet reads. The result is a single deeply-nested JSON object consumed directly by the Vue profile view and D3 tree.

---

## Known Limitations

- **Scale** — Google Sheets is suitable for family trees up to a few thousand rows per sheet. Beyond that, the in-memory load-all approach in `getPedigreePayload` may become slow.
- **No offline support** — requires an active Google session.
- **Apps Script quotas** — Google enforces daily limits on script executions and spreadsheet API calls. For typical single-family use these limits are not a concern. See [Apps Script quotas](https://developers.google.com/apps-script/guides/services/quotas) for details.
- **No automated tests** — concurrent-write correctness relies on `LockService` and has not been verified under load.

---

## Schema Reference

See [`schema.md`](schema.md) for the complete column-level schema, foreign key diagram, and the Apps Script pattern for assembling flat rows into nested JSON.

---

## License

MIT License

Copyright (c) 2026

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
