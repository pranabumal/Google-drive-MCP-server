# Google Drive MCP reader

This workspace MCP server provides read-only Google Drive search, file metadata, and text extraction for Google Docs, Google Sheets, Word `.docx`, Excel `.xlsx`, CSV, and plain-text files.

## Setup

1. Install Python 3.10 or newer. From this folder, create the workspace environment and install dependencies:

   ```sh
   python3 -m venv ../.venv
   ../.venv/bin/python -m pip install -r requirements.txt
   ```

2. In [Google Cloud Console](https://console.cloud.google.com/apis/credentials), create or select a project, enable the Google Drive API, and create an OAuth client ID with application type **Desktop app**.
3. Set the OAuth client ID and secret in the environment used to launch VS Code. For a terminal session on macOS:

   ```sh
   export GOOGLE_OAUTH_CLIENT_ID='your-client-id.apps.googleusercontent.com'
   export GOOGLE_OAUTH_CLIENT_SECRET='your-client-secret'
   ```

   Do not add OAuth secrets or token files to this project or share them in chat.
4. From this folder, run `../.venv/bin/python scripts/authorize.py` and approve the read-only Google Drive access in the browser. The refresh token is saved with restricted file permissions under `~/.config/panopto-google-drive-mcp/`.
5. Launch VS Code from the workspace root in the same terminal with `cd .. && code .`, then start or restart the `google-drive` MCP server from VS Code's MCP Servers view.

## Tools

- `search_drive_files`: find files by name or indexed content.
- `get_drive_file_details`: retrieve metadata and a Drive link.
- `read_drive_file`: extract document text or spreadsheet rows.

The server does not modify, create, or share Drive files. File reads are limited to 20 MB and tool output is capped at 50,000 characters. Legacy `.doc` and `.xls`, PDFs, and other unsupported formats are not extracted.
