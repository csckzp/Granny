#!/usr/bin/env python3
"""
deploy.py
---------
Automates the full Granny setup in one command:
  1. Builds the 6-tab spreadsheet and uploads it to Google Drive as a Google Sheet
  2. Creates a container-bound Apps Script project
  3. Uploads Code.gs, Index.html, and the appsscript.json manifest

After this script finishes you still need to do two things manually:
  A. Open the spreadsheet URL printed below and add yourself to the Users tab (row 2).
  B. Open the Apps Script URL, then Deploy → New deployment → Web app
       Execute as: Me
       Who has access: Anyone with a Google account
     Copy the resulting /exec URL — that is your live app URL.

--- One-time Google Cloud setup ---

  1. Go to https://console.cloud.google.com and create (or select) a project.
  2. Enable both APIs:
       APIs & Services → Library → search "Google Drive API" → Enable
       APIs & Services → Library → search "Apps Script API" → Enable
  3. Create OAuth credentials:
       APIs & Services → Credentials → Create credentials → OAuth client ID
       Application type: Desktop app → Create → Download JSON
       Rename the downloaded file to credentials.json and place it here.
  4. Enable the Apps Script API for your Google account (one-time per account):
       https://script.google.com/home/usersettings → Google Apps Script API → On

--- Python dependencies ---

  pip install openpyxl google-api-python-client google-auth-oauthlib
"""

import io
import json
import os
import sys

from create_template import build_workbook

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ---------------------------------------------------------------------------
# OAuth scopes
# ---------------------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/drive.file",    # create/edit files this app owns
    "https://www.googleapis.com/auth/script.projects",  # create/edit Apps Script projects
]

CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")
TOKEN_FILE       = os.path.join(os.path.dirname(__file__), "token.json")

# appsscript.json manifest pushed into every new project.
# Change timeZone to match your locale if you use time-based triggers.
APPSSCRIPT_MANIFEST = {
    "timeZone": "America/New_York",
    "dependencies": {},
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8",
    "webapp": {
        "executeAs": "USER_DEPLOYING",
        "access": "ANYONE_WITH_GOOGLE_ACCOUNT",
    },
}


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def get_credentials() -> Credentials:
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                sys.exit(
                    "ERROR: credentials.json not found.\n"
                    "Follow the 'One-time Google Cloud setup' steps in this file's docstring."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as fh:
            fh.write(creds.to_json())

    return creds


# ---------------------------------------------------------------------------
# Step 1 — upload spreadsheet
# ---------------------------------------------------------------------------

def upload_spreadsheet(drive_service) -> tuple[str, str]:
    """
    Build the workbook in memory, upload it to Drive, and convert it to
    Google Sheets format in a single API call.

    Returns (spreadsheet_id, web_view_link).
    """
    print("Building spreadsheet template...")
    wb = build_workbook()

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    print("Uploading to Google Drive and converting to Google Sheets...")
    media = MediaIoBaseUpload(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False,
    )
    file_meta = {
        "name": "Granny",
        "mimeType": "application/vnd.google-apps.spreadsheet",
    }
    result = drive_service.files().create(
        body=file_meta,
        media_body=media,
        fields="id,webViewLink",
    ).execute()

    return result["id"], result["webViewLink"]


# ---------------------------------------------------------------------------
# Step 2 — create Apps Script project bound to the spreadsheet
# ---------------------------------------------------------------------------

def create_script_project(script_service, spreadsheet_id: str) -> tuple[str, str]:
    """
    Create a container-bound Apps Script project.

    Returns (script_id, script_editor_url).
    """
    print("Creating Apps Script project...")
    project = script_service.projects().create(body={
        "title": "Granny",
        "parentId": spreadsheet_id,
    }).execute()

    script_id = project["scriptId"]
    editor_url = f"https://script.google.com/d/{script_id}/edit"
    return script_id, editor_url


# ---------------------------------------------------------------------------
# Step 3 — upload project files
# ---------------------------------------------------------------------------

def _read_local(filename: str) -> str | None:
    path = os.path.join(os.path.dirname(__file__), filename)
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as fh:
        return fh.read()


def upload_script_files(script_service, script_id: str) -> None:
    """
    Push appsscript.json, Code.gs, and Index.html to the project.
    The PUT /content call replaces all files, so we always include the manifest.
    """
    print("Uploading script files...")

    files = [
        {
            "name": "appsscript",
            "type": "JSON",
            "source": json.dumps(APPSSCRIPT_MANIFEST, indent=2),
        }
    ]

    code_gs = _read_local("Code.gs")
    if code_gs:
        files.append({"name": "Code", "type": "SERVER_JS", "source": code_gs})
        print("  [OK] Code.gs")
    else:
        print("  [SKIP] Code.gs not found — add it manually in the Apps Script editor")

    index_html = _read_local("Index.html")
    if index_html:
        files.append({"name": "Index", "type": "HTML", "source": index_html})
        print("  [OK] Index.html")
    else:
        print("  [SKIP] Index.html not found — add it manually in the Apps Script editor")

    script_service.projects().updateContent(
        scriptId=script_id,
        body={"files": files},
    ).execute()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    creds          = get_credentials()
    drive_service  = build("drive",  "v3", credentials=creds)
    script_service = build("script", "v1", credentials=creds)

    spreadsheet_id, sheet_url  = upload_spreadsheet(drive_service)
    script_id,      editor_url = create_script_project(script_service, spreadsheet_id)
    upload_script_files(script_service, script_id)

    print()
    print("=" * 60)
    print("Setup complete.")
    print()
    print(f"  Spreadsheet : {sheet_url}")
    print(f"  Script editor: {editor_url}")
    print()
    print("Remaining manual steps:")
    print()
    print("  1. Open the spreadsheet, click the 'Users' tab, and add")
    print("     yourself in row 2:")
    print("       user_id     USR0001")
    print("       email       <your Google account email>")
    print("       display_name  <your name>")
    print("       role        admin")
    print("       created_at  <today's date>")
    print()
    print("  2. Open the script editor (URL above), then:")
    print("       Deploy → New deployment → gear icon → Web app")
    print("       Execute as: Me")
    print("       Who has access: Anyone with a Google account")
    print("       → Deploy → copy the /exec URL")
    print()
    print("  That /exec URL is your live app.")
    print("=" * 60)


if __name__ == "__main__":
    main()
