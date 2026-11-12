"""Small, dependency-free security helpers shared across the server.

Each helper addresses one class of untrusted input:
- email content (subjects, bodies, senders) that ends up in exported files
- caller-supplied file names for exports
- URLs returned by remote APIs that would receive a bearer token
- connection strings that may embed credentials
"""

import os
import re
from typing import Optional
from urllib.parse import quote, urlparse

# Cells starting with these characters are interpreted as formulas by
# Excel / LibreOffice / Google Sheets (CSV injection, OWASP).
_CSV_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")

# Only Microsoft Graph may receive our Graph bearer token.
TRUSTED_GRAPH_HOSTS = frozenset({"graph.microsoft.com"})

_WINDOWS_RESERVED_NAMES = frozenset(
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{i}" for i in range(1, 10)}
    | {f"LPT{i}" for i in range(1, 10)}
)

_URI_CREDENTIALS_RE = re.compile(r"(?P<scheme>[a-zA-Z][a-zA-Z0-9+.\-]*://)[^/\s@]+@")


class UnsafePathError(ValueError):
    """Raised when a caller-supplied export path is not allowed."""


def csv_safe(value):
    """Neutralise spreadsheet formula injection in a single CSV cell."""
    if isinstance(value, str) and value.startswith(_CSV_FORMULA_PREFIXES):
        return "'" + value
    return value


def resolve_export_path(export_dir: str, requested_name: Optional[str], default_name: str,
                        extension: str) -> str:
    """Return a path for a new export file that is guaranteed to live in ``export_dir``.

    ``requested_name`` may only be a bare file name: directory components,
    absolute paths, drive letters and ``..`` are rejected rather than silently
    stripped, so callers notice. The extension is forced to ``extension``.
    """
    name = requested_name.strip() if requested_name else default_name
    if not name:
        raise UnsafePathError("Empty export file name")

    if (
        os.path.isabs(name)
        or "/" in name
        or "\\" in name
        or ":" in name  # drive letters and NTFS alternate data streams
        or name in (".", "..")
        or "\x00" in name
    ):
        raise UnsafePathError(
            "output_path must be a bare file name; exports are always written to the exports directory"
        )

    if name.split(".")[0].strip().upper() in _WINDOWS_RESERVED_NAMES:
        raise UnsafePathError("Reserved device name is not allowed as an export file name")

    stem, ext = os.path.splitext(name)
    if ext.lower() != f".{extension}":
        name = f"{stem or name}.{extension}"

    export_root = os.path.realpath(export_dir)
    candidate = os.path.realpath(os.path.join(export_root, name))
    if os.path.dirname(candidate) != export_root:
        raise UnsafePathError("Export path escapes the exports directory")
    return candidate


def is_trusted_graph_url(url: str) -> bool:
    """True only for https URLs on a Microsoft Graph host (used for @odata.nextLink)."""
    try:
        parsed = urlparse(url)
    except ValueError:
        return False
    return (
        parsed.scheme == "https"
        and (parsed.hostname or "").lower() in TRUSTED_GRAPH_HOSTS
        and parsed.port in (None, 443)
        and not parsed.username
        and not parsed.password
    )


def quote_path_segment(value: str) -> str:
    """Percent-encode a value for use as one URL path segment (keeps '@')."""
    return quote(value, safe="@")


def redact_uri(text: Optional[str]) -> str:
    """Strip ``user:password@`` credentials from any URIs inside ``text``."""
    if not text:
        return ""
    return _URI_CREDENTIALS_RE.sub(r"\g<scheme>***@", str(text))


def clamp(value: int, low: int, high: int) -> int:
    """Clamp an integer into [low, high]; non-integers fall back to ``low``."""
    try:
        value = int(value)
    except (TypeError, ValueError):
        return low
    return max(low, min(high, value))
