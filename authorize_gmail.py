"""Run locally once to authorize Gmail and generate a valid token.json.

Keep credentials.json and token.json private. Never commit either file to GitHub.
"""
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow

ROOT = Path(__file__).resolve().parent
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
credentials_file = ROOT / "credentials.json"
token_file = ROOT / "token.json"

if not credentials_file.exists():
    raise SystemExit(
        "Missing credentials.json. Download an OAuth Desktop app client JSON from Google Cloud, "
        "save it beside this script as credentials.json, and rerun."
    )

flow = InstalledAppFlow.from_client_secrets_file(str(credentials_file), SCOPES)
creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")
token_file.write_text(creds.to_json(), encoding="utf-8")
print(f"Valid OAuth token JSON generated at: {token_file}")
print("Keep token.json private. Do not upload it or commit it to source control.")
