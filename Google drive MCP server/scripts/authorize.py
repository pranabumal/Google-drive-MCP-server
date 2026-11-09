from __future__ import annotations

import os
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
TOKEN_PATH = Path(
    os.environ.get(
        "GOOGLE_DRIVE_TOKEN_PATH",
        Path.home() / ".config" / "panopto-google-drive-mcp" / "token.json",
    )
)


def main() -> None:
    client_id = os.environ.get("GOOGLE_OAUTH_CLIENT_ID")
    client_secret = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET")
    if not client_id or not client_secret:
        raise SystemExit(
            "Set GOOGLE_OAUTH_CLIENT_ID and GOOGLE_OAUTH_CLIENT_SECRET before authorizing."
        )

    client_config = {
        "installed": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "redirect_uris": ["http://localhost"],
        }
    }
    flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
    credentials = flow.run_local_server(host="127.0.0.1", port=0, open_browser=True)

    TOKEN_PATH.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    TOKEN_PATH.parent.chmod(0o700)
    TOKEN_PATH.write_text(credentials.to_json(), encoding="utf-8")
    TOKEN_PATH.chmod(0o600)
    print(f"Google Drive read-only authorization saved to {TOKEN_PATH}")


if __name__ == "__main__":
    main()
