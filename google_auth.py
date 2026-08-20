"""
One-time Google OAuth setup.
Run once: python google_auth.py
Saves token.json next to this script for use by gdocs.py.
"""
import json
import ssl
import urllib.parse
import urllib.request
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path

_HERE      = Path(__file__).parent
CREDS_FILE = _HERE / "client_secret.json"   # rename your downloaded credentials file to this
TOKEN_FILE = _HERE / "token.json"

SCOPES       = ["https://www.googleapis.com/auth/drive.file"]
REDIRECT_URI = "http://localhost:8765"

_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode    = ssl.CERT_NONE

if not CREDS_FILE.exists():
    print(f"ERROR: Put your Google OAuth credentials JSON at:\n  {CREDS_FILE}")
    print("Download it from Google Cloud Console → APIs & Services → Credentials → OAuth 2.0 Client IDs")
    raise SystemExit(1)

with open(CREDS_FILE) as f:
    raw = json.load(f)
creds         = raw.get("installed") or raw.get("web")
CLIENT_ID     = creds["client_id"]
CLIENT_SECRET = creds["client_secret"]
TOKEN_URI     = creds["token_uri"]

_auth_code = None

class _Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        global _auth_code
        params = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        if "code" in params:
            _auth_code = params["code"][0]
            self.send_response(200); self.end_headers()
            self.wfile.write(b"<h2>Authentication successful! You can close this tab.</h2>")
        else:
            self.send_response(400); self.end_headers()
            self.wfile.write(b"<h2>No code received.</h2>")
    def log_message(self, *a): pass

auth_url = ("https://accounts.google.com/o/oauth2/auth?" + urllib.parse.urlencode({
    "client_id": CLIENT_ID, "redirect_uri": REDIRECT_URI,
    "response_type": "code", "scope": " ".join(SCOPES),
    "access_type": "offline", "prompt": "consent",
}))

print("Opening browser for Google sign-in...")
webbrowser.open(auth_url)

srv = HTTPServer(("localhost", 8765), _Handler)
srv.timeout = 120
print("Waiting for redirect (timeout 120s)...")
while not _auth_code:
    srv.handle_request()
srv.server_close()

data = urllib.parse.urlencode({
    "code": _auth_code, "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET, "redirect_uri": REDIRECT_URI,
    "grant_type": "authorization_code",
}).encode()
req = urllib.request.Request(TOKEN_URI, data=data, method="POST")
req.add_header("Content-Type", "application/x-www-form-urlencoded")
with urllib.request.urlopen(req, context=_SSL, timeout=30) as resp:
    token_data = json.loads(resp.read())

token_data["client_id"]     = CLIENT_ID
token_data["client_secret"] = CLIENT_SECRET
with open(TOKEN_FILE, "w") as f:
    json.dump(token_data, f, indent=2)

print(f"\nSUCCESS — token saved to: {TOKEN_FILE}")
print("You can now use 'Generate LP' and get a Google Docs link.")
