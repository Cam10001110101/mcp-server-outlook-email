"""Regression tests for the security hardening (exports, Graph URLs, AppleScript, redaction)."""

import csv
import json
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from exporters import render_html, write_email_export  # noqa: E402
from security_utils import (  # noqa: E402
    UnsafePathError,
    clamp,
    csv_safe,
    is_trusted_graph_url,
    quote_path_segment,
    redact_uri,
    resolve_export_path,
)

EVIL_EMAIL = {
    "id": "1",
    "Subject": "<script>alert(1)</script>",
    "SenderEmailAddress": '"><img src=x onerror=alert(2)>@evil.test',
    "ReceivedTime": "2026-01-01T00:00:00",
    "Folder": "Inbox",
    "Body": "=HYPERLINK(\"http://evil.test\",\"click\") <iframe src=//evil.test>",
}


# --- export path confinement ------------------------------------------------

@pytest.mark.parametrize(
    "name",
    [
        "../escape.json",
        "../../etc/passwd",
        "/etc/passwd",
        "sub/dir.json",
        "..\\win.json",
        "C:evil.json",
        "C:\\Windows\\evil.json",
        "..",
        ".",
        "a\x00b.json",
    ],
)
def test_resolve_export_path_rejects_non_bare_names(tmp_path, name):
    with pytest.raises(UnsafePathError):
        resolve_export_path(str(tmp_path), name, "default", "json")


def test_resolve_export_path_forces_extension_and_stays_inside(tmp_path):
    path = resolve_export_path(str(tmp_path), "report.sh", "default", "csv")
    assert path == os.path.join(os.path.realpath(tmp_path), "report.csv")
    assert resolve_export_path(str(tmp_path), "report.CSV", "d", "csv").endswith("report.CSV")
    assert resolve_export_path(str(tmp_path), None, "default", "json").endswith("default.json")


def test_resolve_export_path_rejects_symlink_escape(tmp_path):
    outside = tmp_path / "outside"
    outside.mkdir()
    exports = tmp_path / "exports"
    exports.mkdir()
    (exports / "link.json").symlink_to(outside / "target.json")
    with pytest.raises(UnsafePathError):
        resolve_export_path(str(exports), "link.json", "d", "json")


def test_write_export_never_overwrites(tmp_path):
    victim = tmp_path / "keep.json"
    victim.write_text("original")
    with pytest.raises(FileExistsError):
        write_email_export([EVIL_EMAIL], "json", str(tmp_path), "keep.json")
    assert victim.read_text() == "original"


def test_write_export_json_roundtrip(tmp_path):
    path, size = write_email_export([EVIL_EMAIL], "json", str(tmp_path / "exports"))
    assert os.path.dirname(path) == os.path.realpath(tmp_path / "exports")
    assert size == os.path.getsize(path)
    assert json.load(open(path, encoding="utf-8"))[0]["Subject"] == EVIL_EMAIL["Subject"]


# --- HTML export escaping ---------------------------------------------------

def test_html_export_escapes_email_fields():
    out = render_html([EVIL_EMAIL])
    assert "<script>alert(1)</script>" not in out
    assert "<img src=x" not in out
    assert "<iframe" not in out
    assert "&lt;script&gt;alert(1)&lt;/script&gt;" in out
    assert "Content-Security-Policy" in out


def test_html_export_file(tmp_path):
    path, _ = write_email_export([EVIL_EMAIL], "html", str(tmp_path))
    assert path.endswith(".html")
    assert "<script>alert" not in open(path, encoding="utf-8").read()


# --- CSV formula injection --------------------------------------------------

@pytest.mark.parametrize("value", ["=1+1", "+1", "-1", "@SUM(A1)", "\tx", "\rx"])
def test_csv_safe_prefixes_formulas(value):
    assert csv_safe(value) == "'" + value


@pytest.mark.parametrize("value", ["hello", "", 5, None, "a=b"])
def test_csv_safe_leaves_plain_values(value):
    assert csv_safe(value) == value


def test_csv_export_neutralises_formulas(tmp_path):
    path, _ = write_email_export([EVIL_EMAIL], "csv", str(tmp_path))
    rows = list(csv.DictReader(open(path, encoding="utf-8", newline="")))
    assert rows[0]["Body"].startswith("'=HYPERLINK")


# --- Graph URL safety -------------------------------------------------------

@pytest.mark.parametrize(
    "url,ok",
    [
        ("https://graph.microsoft.com/v1.0/users?$skiptoken=abc", True),
        ("https://GRAPH.microsoft.com/v1.0/me/messages", True),
        ("http://graph.microsoft.com/v1.0/users", False),
        ("https://graph.microsoft.com.evil.test/v1.0", False),
        ("https://evil.test/graph.microsoft.com", False),
        ("https://user:pw@graph.microsoft.com/v1.0", False),
        ("https://graph.microsoft.com:8443/v1.0", False),
        ("https://evil.test@graph.microsoft.com/", False),
        ("not a url", False),
    ],
)
def test_is_trusted_graph_url(url, ok):
    assert is_trusted_graph_url(url) is ok


def test_quote_path_segment_blocks_path_and_query_injection():
    assert quote_path_segment("user@contoso.com") == "user@contoso.com"
    encoded = quote_path_segment("a@b.com/../../admin?x=1#y")
    assert "/" not in encoded and "?" not in encoded and "#" not in encoded


def test_graph_endpoint_is_encoded():
    pytest.importorskip("requests")
    from connectors.graph_connector import GraphAPIConnector

    conn = GraphAPIConnector(client_id="x", client_secret="y", tenant_id="t",
                             user_emails=["a@b.com"])
    assert conn._get_user_endpoint("a@b.com/../me") == "/users/a@b.com%2F..%2Fme"
    assert conn._get_folder_id("Inbox") == "inbox"
    assert "/" not in conn._get_folder_id("x/../y")


# --- AppleScript folder whitelist -------------------------------------------

def test_mac_folder_whitelist():
    from connectors.mac_connector import MacOutlookConnector

    conn = MacOutlookConnector()
    assert conn._get_folder_script_name("Inbox", "exchange") == "inbox"
    with pytest.raises(ValueError):
        conn._get_folder_script_name('inbox\ndo shell script "id"', "exchange")


# --- misc -------------------------------------------------------------------

@pytest.mark.parametrize(
    "text,expected",
    [
        ("mongodb://u:p@host:27017/db", "mongodb://***@host:27017/db"),
        ("mongodb+srv://u:p@cluster.x.net/db", "mongodb+srv://***@cluster.x.net/db"),
        ("mongodb://host:27017/db", "mongodb://host:27017/db"),
        ("error talking to mongodb://u:p@h/db now", "error talking to mongodb://***@h/db now"),
        (None, ""),
    ],
)
def test_redact_uri(text, expected):
    assert redact_uri(text) == expected


def test_clamp():
    assert clamp(10_000, 1, 200) == 200
    assert clamp(-5, 1, 200) == 1
    assert clamp("x", 1, 200) == 1
    assert clamp(20, 1, 200) == 20
