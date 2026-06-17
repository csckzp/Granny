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
| `create_template.py` | Python script that generates the pre-configured `granny_template.xlsx` |
| `schema.md` | Full column-level schema reference with FK diagram and JSON assembly notes |

---

## Deployment

### Prerequisites

- A Google account
- Access to [Google Drive](https://drive.google.com) and [Google Sheets](https://sheets.google.com)
- No local tools required

---

### Step 1 — Create the Google Spreadsheet

The easiest way is to generate the template file with the included Python script, then convert it in Google Drive. Manual creation is also documented below if you prefer.

#### Option A — Generate from the Python template (recommended)

1. Make sure Python 3 and `openpyxl` are installed:
   ```bash
   pip install openpyxl
   ```
2. Run the script:
   ```bash
   python create_template.py
   # → Saved granny_template.xlsx
   ```
3. Upload the file to Google Drive:
   **Drive → New → File upload → select `granny_template.xlsx`**
4. Right-click the uploaded file → **Open with → Google Sheets**.
5. Inside Google Sheets: **File → Save as Google Sheets**.
6. Name the new sheet **Granny** (or any name you prefer) and click **OK**.

The resulting sheet has all six tabs, bold indigo headers, frozen header rows, column widths, and dropdown validation — ready to go.

---

#### Option B — Create manually

1. Go to [sheets.google.com](https://sheets.google.com) and create a **Blank** spreadsheet.
2. Name it **Granny** (or any name you prefer).
3. Create six sheet tabs by clicking the **+** button at the bottom.

**Tab names are case-sensitive. Headers must be in row 1.**

##### Tab 1: `Users`
```
user_id | email | display_name | role | created_at
```

##### Tab 2: `People`
```
person_id | first_name | middle_name | last_name | sex | birth_date | birth_place | death_date | death_place | is_living | notes | created_by | created_at | updated_at
```

##### Tab 3: `Families`
```
family_id | spouse1_id | spouse2_id | union_type | union_date | union_place | union_end_date | union_end_reason | notes | created_by | created_at | updated_at
```

##### Tab 4: `Family_Children`
```
family_id | child_id | relationship_type | notes
```

##### Tab 5: `Events`
```
event_id | event_type | event_date | event_place | title | description | source_citation | created_by | created_at | updated_at
```

##### Tab 6: `Event_Participants`
```
event_id | person_id | role | notes
```

> **Tip:** You can delete the default `Sheet1` tab once all six are created.

---

### Step 2 — Open the Apps Script Editor

1. In your spreadsheet, click **Extensions → Apps Script**.
2. The script editor opens in a new tab, bound to your spreadsheet.

---

### Step 3 — Add the Backend (`Code.gs`)

1. In the editor, click on the default file named **Code.gs** in the left sidebar.
2. **Select all** existing content and **delete** it (the default `myFunction` stub).
3. Paste the entire contents of `Code.gs` from this repository.
4. Click the **Save** icon (or press `Ctrl+S` / `Cmd+S`).

---

### Step 4 — Add the Frontend (`Index.html`)

1. In the editor, click the **+** button next to "Files" in the left sidebar and choose **HTML**.
2. Name the file **`Index`** (exactly — no extension; Apps Script appends `.html` automatically).
3. **Select all** default content and **delete** it.
4. Paste the entire contents of `Index.html` from this repository.
5. Save the file.

Your file list in the editor should now show:
```
Code.gs
Index.html
```

---

### Step 5 — Register Yourself as an Admin

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

### Step 6 — Deploy as a Web App

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

### Step 7 — Open the App

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
