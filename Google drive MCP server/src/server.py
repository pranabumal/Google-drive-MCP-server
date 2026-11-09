from __future__ import annotations

import io
import os
from pathlib import Path
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from mcp.server.fastmcp import FastMCP
from openpyxl import load_workbook
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
DRIVE_API = "drive"
MAX_FILE_BYTES = 20 * 1024 * 1024
MAX_OUTPUT_CHARACTERS = 50_000
MAX_SPREADSHEET_ROWS = 3_000
MAX_SPREADSHEET_COLUMNS = 200
TOKEN_PATH = Path(
    os.environ.get(
        "GOOGLE_DRIVE_TOKEN_PATH",
        Path.home() / ".config" / "panopto-google-drive-mcp" / "token.json",
    )
)

mcp = FastMCP("google-drive-reader")


def _client_credentials() -> tuple[str, str]:
    client_id = os.environ.get("GOOGLE_OAUTH_CLIENT_ID")
    client_secret = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET")
    if not client_id or not client_secret:
        raise RuntimeError(
            "Google OAuth credentials are not configured. Set "
            "GOOGLE_OAUTH_CLIENT_ID and GOOGLE_OAUTH_CLIENT_SECRET, "
            "then run `python scripts/authorize.py`."
        )
    return client_id, client_secret


def _save_token(credentials: Credentials) -> None:
    TOKEN_PATH.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    TOKEN_PATH.parent.chmod(0o700)
    TOKEN_PATH.write_text(credentials.to_json(), encoding="utf-8")
    TOKEN_PATH.chmod(0o600)


def _drive_service() -> Any:
    client_id, _ = _client_credentials()
    if not TOKEN_PATH.exists():
        raise RuntimeError(
            "Google Drive is not authorized yet. Set the OAuth environment "
            "variables and run `python scripts/authorize.py` from this workspace."
        )

    try:
        credentials = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    except (OSError, ValueError) as error:
        raise RuntimeError(f"Could not load Google OAuth token: {error}") from error

    if not credentials.valid:
        if not credentials.expired or not credentials.refresh_token:
            raise RuntimeError(
                "Google authorization has expired or was revoked. Run "
                "`python scripts/authorize.py` again."
            )
        credentials.refresh(Request())
        _save_token(credentials)

    return build(DRIVE_API, "v3", credentials=credentials, cache_discovery=False)


def _file_metadata(service: Any, file_id: str) -> dict[str, Any]:
    return (
        service.files()
        .get(
            fileId=file_id,
            fields="id,name,mimeType,size,modifiedTime,webViewLink,description,owners(displayName,emailAddress)",
            supportsAllDrives=True,
        )
        .execute()
    )


def _download(service: Any, file_id: str, export_mime_type: str | None = None) -> bytes:
    if export_mime_type:
        request = service.files().export_media(fileId=file_id, mimeType=export_mime_type)
    else:
        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)

    target = io.BytesIO()
    downloader = MediaIoBaseDownload(target, request, chunksize=256 * 1024)
    done = False
    while not done:
        _, done = downloader.next_chunk()
        if target.tell() > MAX_FILE_BYTES:
            raise ValueError("This file is larger than the 20 MB reading limit.")
    return target.getvalue()


def _spreadsheet_text(contents: bytes) -> str:
    workbook = load_workbook(io.BytesIO(contents), read_only=True, data_only=True)
    sections: list[str] = []
    rows_read = 0
    truncated = False

    try:
        for worksheet in workbook.worksheets:
            lines = [f"## {worksheet.title}"]
            if (worksheet.max_row or 0) > MAX_SPREADSHEET_ROWS:
                truncated = True
            for row in worksheet.iter_rows(
                max_row=MAX_SPREADSHEET_ROWS,
                max_col=MAX_SPREADSHEET_COLUMNS,
                values_only=True,
            ):
                if rows_read >= MAX_SPREADSHEET_ROWS:
                    truncated = True
                    break
                if not any(value is not None for value in row):
                    continue
                cells = ["" if value is None else str(value).replace("\n", " ") for value in row]
                lines.append("\t".join(cells))
                rows_read += 1
            sections.append("\n".join(lines))
            if rows_read >= MAX_SPREADSHEET_ROWS:
                truncated = True
                break
        if truncated:
            sections.append(
                "[Output capped at 3,000 rows across the workbook; some content may be omitted.]"
            )
    finally:
        workbook.close()
    return "\n\n".join(sections)


def _word_text(contents: bytes) -> str:
    document = Document(io.BytesIO(contents))
    sections = []
    for block in document.iter_inner_content():
        if isinstance(block, Paragraph):
            if block.text:
                sections.append(block.text)
        elif isinstance(block, Table):
            sections.extend(
                "\t".join(cell.text.replace("\n", " ") for cell in row.cells)
                for row in block.rows
            )
    return "\n".join(sections)


def _extract_file(service: Any, file: dict[str, Any]) -> str:
    name = file.get("name", "")
    extension = Path(name).suffix.lower()
    mime_type = file.get("mimeType", "")

    if mime_type == "application/vnd.google-apps.document":
        return _download(service, file["id"], "text/plain").decode("utf-8")
    if mime_type == "application/vnd.google-apps.spreadsheet":
        exported = _download(
            service,
            file["id"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        return _spreadsheet_text(exported)
    if (
        mime_type
        == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        or extension == ".docx"
    ):
        return _word_text(_download(service, file["id"]))
    if (
        mime_type
        == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        or extension == ".xlsx"
    ):
        return _spreadsheet_text(_download(service, file["id"]))
    if mime_type in {"text/csv", "text/plain"} or extension in {".csv", ".txt"}:
        return _download(service, file["id"]).decode("utf-8-sig")

    raise ValueError(
        f"Reading {mime_type or extension or 'this file type'} is not supported. "
        "Supported formats are Google Docs, Google Sheets, .docx, .xlsx, .csv, and .txt."
    )


@mcp.tool()
def search_drive_files(query: str, limit: int = 10) -> list[dict[str, Any]]:
    """Search Drive filenames and indexed file contents; use a result's ID to read its file."""
    if not query.strip():
        raise ValueError("Search query must not be empty.")
    if len(query) > 200:
        raise ValueError("Search query must be 200 characters or fewer.")
    if not 1 <= limit <= 50:
        raise ValueError("Limit must be between 1 and 50.")

    escaped = query.replace("\\", "\\\\").replace("'", "\\'")
    service = _drive_service()
    result = (
        service.files()
        .list(
            q=f"trashed = false and (name contains '{escaped}' or fullText contains '{escaped}')",
            pageSize=limit,
            orderBy="modifiedTime desc",
            fields="files(id,name,mimeType,size,modifiedTime,webViewLink,description)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        )
        .execute()
    )
    return result.get("files", [])


@mcp.tool()
def get_drive_file_details(file_id: str) -> dict[str, Any]:
    """Get metadata, owner, modified date, and browser link for a Drive file."""
    if not file_id.strip():
        raise ValueError("File ID must not be empty.")
    return _file_metadata(_drive_service(), file_id)


@mcp.tool()
def read_drive_file(file_id: str, max_characters: int = 30_000) -> dict[str, Any]:
    """Extract text from a Google Doc or Sheet, Word .docx, Excel .xlsx, CSV, or text file."""
    if not file_id.strip():
        raise ValueError("File ID must not be empty.")
    if not 1 <= max_characters <= MAX_OUTPUT_CHARACTERS:
        raise ValueError(f"max_characters must be between 1 and {MAX_OUTPUT_CHARACTERS}.")

    service = _drive_service()
    file = _file_metadata(service, file_id)
    content = _extract_file(service, file)
    return {
        "file": {
            "id": file.get("id"),
            "name": file.get("name"),
            "mimeType": file.get("mimeType"),
            "modifiedTime": file.get("modifiedTime"),
            "webViewLink": file.get("webViewLink"),
        },
        "content": content[:max_characters],
        "truncated": len(content) > max_characters,
        "returnedCharacters": min(len(content), max_characters),
        "totalCharacters": len(content),
    }


if __name__ == "__main__":
    mcp.run(transport="stdio")
