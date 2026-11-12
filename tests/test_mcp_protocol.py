"""Wire-level MCP 2026-07-28 checks, without mail, credentials or external services."""

import asyncio
import json
import os
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from jsonschema import Draft202012Validator
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from starlette.testclient import TestClient

sys.path.insert(0, str(Path(__file__).parents[1] / "src"))
import mcp_server as app  # noqa: E402

VERSION = "2026-07-28"
PREFIX = "io.modelcontextprotocol/"


def request(method, params=None, *, version=VERSION, request_id=1):
    return {
        "jsonrpc": "2.0", "id": request_id, "method": method,
        "params": {
            "_meta": {
                PREFIX + "protocolVersion": version,
                PREFIX + "clientCapabilities": {},
                PREFIX + "clientInfo": {"name": "protocol-test", "version": "1"},
            },
            **(params or {}),
        },
    }


def headers(message):
    values = {
        "Accept": "application/json, text/event-stream",
        "MCP-Protocol-Version": message["params"]["_meta"][PREFIX + "protocolVersion"],
        "Mcp-Method": message["method"],
    }
    if "name" in message["params"]:
        values["Mcp-Name"] = message["params"]["name"]
    elif "uri" in message["params"]:
        values["Mcp-Name"] = message["params"]["uri"]
    return values


def response_messages(response):
    if response.headers["content-type"].startswith("application/json"):
        return [response.json()]
    assert response.headers["content-type"].startswith("text/event-stream")
    assert "\nid:" not in response.text
    return [json.loads(line[6:]) for line in response.text.splitlines() if line.startswith("data: ")]


@pytest.fixture
def http():
    application = app.mcp.streamable_http_app(
        stateless_http=True, transport_security=app.local_http_security(),
    )
    with TestClient(application, base_url="http://127.0.0.1:8000") as client:
        yield client


@pytest.fixture(params=["stdio", "http"])
async def peer(request, tmp_path):
    if request.param == "http":
        application = app.mcp.streamable_http_app(
            stateless_http=True, transport_security=app.local_http_security(),
        )
        with TestClient(application, base_url="http://127.0.0.1:8000") as client:
            async def exchange(message):
                response = client.post("/mcp", json=message, headers=headers(message))
                assert "mcp-session-id" not in response.headers
                return response_messages(response)[-1]
            yield exchange
        return

    # Launch the real CLI with no application configuration or inherited credentials.
    env = {k: v for k, v in os.environ.items() if k in {"PATH", "SYSTEMROOT", "WINDIR", "TEMP", "TMP"}}
    env["PYTHONUTF8"] = "1"
    with (tmp_path / "stderr.log").open("wb") as stderr:
        process = await asyncio.create_subprocess_exec(
            sys.executable, str(Path(app.__file__).resolve()), cwd=tmp_path, env=env,
            stdin=asyncio.subprocess.PIPE, stdout=asyncio.subprocess.PIPE, stderr=stderr,
        )
        async def exchange(message):
            process.stdin.write(json.dumps(message).encode() + b"\n")
            await process.stdin.drain()
            line = await asyncio.wait_for(process.stdout.readline(), timeout=45)
            assert line, (tmp_path / "stderr.log").read_text()
            response = json.loads(line)
            assert response["id"] == message["id"]
            return response
        try:
            yield exchange
        finally:
            process.stdin.close()
            try:
                await asyncio.wait_for(process.wait(), timeout=10)
            except TimeoutError:
                process.kill()
                await process.wait()
                pytest.fail("Server failed to shut down after stdin EOF")


def complete(response, *, cacheable=False):
    assert response["jsonrpc"] == "2.0"
    result = response["result"]
    assert result["resultType"] == "complete"
    assert result["_meta"][PREFIX + "serverInfo"] == {"name": "outlook-email", "version": "0.2.0"}
    if cacheable:
        assert result["ttlMs"] == 0
        assert result["cacheScope"] == "private"
    return result


async def test_native_discovery_and_catalogs(peer):
    # tools/list is valid as the very first request: no initialization or discovery.
    tools = complete(await peer(request("tools/list")), cacheable=True)["tools"]
    assert len(tools) == 11
    assert tools == complete(await peer(request("tools/list", request_id=2)), cacheable=True)["tools"]
    discovered = complete(await peer(request("server/discover", request_id=3)), cacheable=True)
    assert VERSION in discovered["supportedVersions"]
    assert set(discovered["capabilities"]) == {"tools", "resources", "prompts"}
    for tool in tools:
        assert "ctx" not in tool["inputSchema"].get("properties", {})
        Draft202012Validator.check_schema(tool["inputSchema"])
        Draft202012Validator.check_schema(tool["outputSchema"])
    for method, key, count in [
        ("prompts/list", "prompts", 5), ("resources/list", "resources", 1),
        ("resources/templates/list", "resourceTemplates", 3),
    ]:
        result = complete(await peer(request(method)), cacheable=True)
        assert len(result[key]) == count
    assert app.processor is None


async def test_request_scoped_versions_and_protocol_errors(peer):
    error = (await peer(request("tools/list", version="2099-01-01")))["error"]
    assert error["code"] == -32022
    assert VERSION in error["data"]["supported"]
    assert error["data"]["requested"] == "2099-01-01"
    complete(await peer(request("tools/list", request_id=2)), cacheable=True)
    for method, params, code in [
        ("does/not/exist", {}, -32601),
        ("resources/read", {"uri": "email://missing"}, -32602),
        ("resources/read", {"uri": "email://recent/not-a-number"}, -32602),
        ("resources/read", {"uri": "email://recent/0"}, -32602),
    ]:
        response = await peer(request(method, params))
        assert response["error"]["code"] == code


async def test_modern_requests_require_client_capability_metadata(peer):
    message = request("tools/list")
    del message["params"]["_meta"][PREFIX + "clientCapabilities"]
    assert "error" in await peer(message)


async def test_legacy_stdio_compatibility():
    params = StdioServerParameters(command=sys.executable, args=[str(Path(app.__file__).resolve())], env={})
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            result = await session.initialize()
            assert result.protocol_version == "2025-11-25"
            assert len((await session.list_tools()).tools) == 11


async def test_structured_errors_and_prompts(peer):
    result = complete(await peer(request("tools/call", {
        "name": "process_emails", "arguments": {"start_date": "bad", "end_date": "bad", "mailboxes": []},
    })))
    assert result["isError"] is True
    assert result["structuredContent"]["success"] is False
    assert json.loads(result["content"][0]["text"]) == result["structuredContent"]
    catalog = complete(await peer(request("tools/list")), cacheable=True)["tools"]
    schema = next(t["outputSchema"] for t in catalog if t["name"] == "process_emails")
    Draft202012Validator(schema).validate(result["structuredContent"])
    invalid = complete(await peer(request("tools/call", {"name": "process_emails", "arguments": {}})))
    assert invalid["isError"] is True
    prompt = complete(await peer(request("prompts/get", {
        "name": "sentiment_report_prompt", "arguments": {"sender_or_topic": "project"},
    })))
    assert prompt["messages"][0]["role"] == "user"
    assert "project" in prompt["messages"][0]["content"]["text"]


@pytest.mark.parametrize("header,value", [
    ("Mcp-Method", "prompts/list"), ("Mcp-Name", "another_tool"),
    ("MCP-Protocol-Version", "2025-11-25"),
])
def test_http_header_mismatch_is_rejected(http, header, value):
    message = request("tools/call", {"name": "process_emails", "arguments": {}})
    response = http.post("/mcp", json=message, headers={**headers(message), header: value})
    assert response.status_code == 400
    assert response.json()["error"]["code"] == -32020


@pytest.mark.parametrize("header", ["Mcp-Method", "Mcp-Name", "MCP-Protocol-Version"])
def test_http_required_headers(http, header):
    message = request("tools/call", {"name": "process_emails", "arguments": {}})
    values = headers(message)
    del values[header]
    assert http.post("/mcp", json=message, headers=values).status_code == 400


def test_http_origin_host_and_batch_rejection(http):
    message = request("tools/list")
    assert http.post("/mcp", json=message, headers={**headers(message), "Origin": "https://attacker.invalid"}).status_code == 403
    assert http.post("/mcp", json=message, headers={**headers(message), "Host": "attacker.invalid"}).status_code == 421
    assert http.post("/mcp", json=[message], headers=headers(message)).status_code == 400
    assert http.get("/mcp", headers={"MCP-Protocol-Version": VERSION}).status_code == 405


def test_http_rejects_duplicate_version_headers(http):
    message = request("tools/list")
    duplicate = list(headers(message).items()) + [("MCP-Protocol-Version", "2025-11-25")]
    response = http.post("/mcp", json=message, headers=duplicate)
    assert response.status_code == 400
    assert response.json()["error"]["code"] == -32020


def test_http_header_guard_has_a_body_limit():
    application = app.mcp.streamable_http_app(stateless_http=True, max_request_body_size=512)
    message = request("tools/list", {"padding": "x" * 1024})
    with TestClient(application, base_url="http://127.0.0.1:8000") as client:
        response = client.post("/mcp", json=message, headers=headers(message))
        assert response.status_code == 413


def test_resource_backend_errors_are_protocol_errors(http, monkeypatch):
    def unavailable():
        raise RuntimeError("synthetic private backend details")
    monkeypatch.setattr(app, "get_processor", unavailable)
    message = request("resources/read", {"uri": "email://statistics/summary"})
    result = response_messages(http.post("/mcp", json=message, headers=headers(message)))[-1]
    assert result["error"]["code"] == -32603
    assert "synthetic private" not in json.dumps(result)


def test_private_resource_and_export_confinement_over_http(http, monkeypatch, tmp_path):
    sqlite = SimpleNamespace(
        db_path=str(tmp_path / "mail.db"),
        get_basic_statistics=lambda: {"total_emails": 1},
        get_emails_for_export=lambda filters: [{"id": "synthetic", "Subject": "Fixture"}],
    )
    monkeypatch.setattr(app, "processor", SimpleNamespace(sqlite=sqlite))
    message = request("resources/read", {"uri": "email://statistics/summary"})
    result = complete(response_messages(http.post("/mcp", json=message, headers=headers(message)))[-1], cacheable=True)
    assert result["contents"][0]["uri"] == "email://statistics/summary"
    for format in ("csv", "html", "json", "excel"):
        message = request("tools/call", {"name": "export_email_data", "arguments": {
            "format": format, "output_path": "../outside.json",
        }})
        result = complete(response_messages(http.post("/mcp", json=message, headers=headers(message)))[-1])
        assert result["isError"] is True
    assert not (tmp_path / "outside.json").exists()


@pytest.mark.parametrize("progress_token", [None, "fixture-progress"])
def test_progress_is_numeric_increasing_and_request_scoped(http, monkeypatch, progress_token):
    processor = object.__new__(app.EmailProcessor)
    processor.outlook = SimpleNamespace(
        get_mailboxes=lambda: ["fixture"], process_deleted_items=False,
        get_emails_within_date_range=lambda *args: [{"id": "synthetic"}],
    )
    processor.sqlite = SimpleNamespace(
        add_or_update_email=lambda email: True,
        get_unprocessed_emails=lambda **kwargs: [{"id": "synthetic"}],
        mark_as_processed=lambda email_id: None,
    )
    processor.embedding_processor = SimpleNamespace(process_batch=lambda batch: (1, 0))
    monkeypatch.setattr(app, "processor", processor)
    message = request("tools/call", {"name": "process_emails", "arguments": {
        "start_date": "2026-01-01", "end_date": "2026-01-02", "mailboxes": ["All"],
    }})
    if progress_token:
        message["params"]["_meta"]["progressToken"] = progress_token
    messages = response_messages(http.post("/mcp", json=message, headers=headers(message)))
    result = complete(messages[-1])
    assert result["isError"] is False
    assert result["structuredContent"]["processed_count"] == 1
    if progress_token:
        progress = [m["params"] for m in messages[:-1]]
        assert progress
        assert all(p["progressToken"] == progress_token and p["total"] == 100 for p in progress)
        assert all(p["progress"] < q["progress"] for p, q in zip(progress, progress[1:]))
        assert progress[-1]["progress"] == 100
    else:
        assert len(messages) == 1
