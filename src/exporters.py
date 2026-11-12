"""Write email exports to disk.

Email fields are untrusted (anyone can send you an email), and the output
file name comes from an MCP tool argument (which an LLM may fill in from
prompt-injected email content). So this module:
- confines every export to ``export_dir`` and never overwrites a file
- HTML-escapes every email-derived value in HTML exports
- neutralises spreadsheet formulas in CSV exports
"""

import csv
import html
import json
import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from security_utils import csv_safe, resolve_export_path

_HTML_HEAD = """<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'">
    <title>Email Export</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .email { border: 1px solid #ccc; margin: 10px 0; padding: 15px; }
        .header { background-color: #f5f5f5; padding: 10px; margin-bottom: 10px; }
        .subject { font-weight: bold; font-size: 1.1em; }
        .meta { color: #666; font-size: 0.9em; }
    </style>
</head>
<body>
    <h1>Email Export Report</h1>
"""


def _esc(value: Any, default: str = "") -> str:
    return html.escape(str(value if value is not None else default), quote=True)


def render_html(emails: List[Dict[str, Any]]) -> str:
    """Render emails as a standalone HTML report with all fields escaped."""
    parts = [_HTML_HEAD]
    # SQLiteHandler returns capitalized field names (Subject, SenderEmailAddress, ...)
    for email in emails:
        body = email.get("Body", "") or ""
        parts.append(
            f"""    <div class="email">
        <div class="header">
            <div class="subject">{_esc(email.get('Subject'), 'No Subject')}</div>
            <div class="meta">From: {_esc(email.get('SenderEmailAddress'), 'Unknown')} | Date: {_esc(email.get('ReceivedTime'), 'Unknown')} | Folder: {_esc(email.get('Folder'), 'Unknown')}</div>
        </div>
        <div class="body">{_esc(body[:500])}{'...' if len(body) > 500 else ''}</div>
    </div>
"""
        )
    parts.append("</body></html>")
    return "".join(parts)


def write_email_export(
    emails: List[Dict[str, Any]],
    fmt: str,
    export_dir: str,
    output_name: Optional[str] = None,
) -> Tuple[str, int]:
    """Write ``emails`` to a new file in ``export_dir``.

    ``fmt`` is "csv", "html" or anything else (written as JSON).
    Returns ``(file_path, size_in_bytes)``. Raises ``UnsafePathError`` for a
    disallowed ``output_name`` and ``FileExistsError`` if the file exists.
    """
    extension = fmt if fmt in ("csv", "html") else "json"
    os.makedirs(export_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = resolve_export_path(
        export_dir, output_name, default_name=f"emails_export_{timestamp}", extension=extension
    )

    # Mode "x": create exclusively; refuses to clobber or follow an existing symlink.
    if extension == "csv":
        with open(file_path, "x", newline="", encoding="utf-8") as f:
            if emails:
                writer = csv.DictWriter(f, fieldnames=list(emails[0].keys()))
                writer.writeheader()
                writer.writerows({k: csv_safe(v) for k, v in row.items()} for row in emails)
    elif extension == "html":
        with open(file_path, "x", encoding="utf-8") as f:
            f.write(render_html(emails))
    else:
        with open(file_path, "x", encoding="utf-8") as f:
            json.dump(emails, f, indent=2, ensure_ascii=False)

    return file_path, os.path.getsize(file_path)
