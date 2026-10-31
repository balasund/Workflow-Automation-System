"""Run once on the computer that hosts the local Streamlit app to authorize Gmail."""
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow

ROOT = Path(__file__).resolve().parent
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
credentials_file = ROOT / "credentials.json"
token_file = ROOT / "token.json"

if not credentials_file.exists():
    raise SystemExit("Place your Google OAuth Desktop client file here as credentials.json, then rerun this script.")

flow = InstalledAppFlow.from_client_secrets_file(str(credentials_file), SCOPES)
creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")
token_file.write_text(creds.to_json(), encoding="utf-8")
print(f"Gmail authorization saved to {token_file}")
print("Keep token.json private. Do not upload it or commit it to source control.")
