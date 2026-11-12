"""Application policies layered on the official MCP protocol implementation."""

import json

from mcp import types
from mcp.server.context import CallNext, HandlerResult, ServerRequestContext
from mcp.server.mcpserver import MCPServer
from mcp.server.transport_security import (
    DEFAULT_MAX_REQUEST_BODY_SIZE,
    RequestBodyLimitMiddleware,
    TransportSecuritySettings,
)
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse


class ModernHeaderGuard(BaseHTTPMiddleware):
    """Close SDK 2.3's header-only routing gap before its legacy HTTP dispatcher."""

    async def dispatch(self, request, call_next):
        if request.method == "POST":
            try:
                body = json.loads(await request.body())
            except (ValueError, RecursionError):
                return await call_next(request)  # SDK produces the parse error.
            params = body.get("params") if isinstance(body, dict) else None
            meta = params.get("_meta") if isinstance(params, dict) else None
            version = meta.get("io.modelcontextprotocol/protocolVersion") if isinstance(meta, dict) else None
            if version is not None and request.headers.getlist("mcp-protocol-version") != [version]:
                return JSONResponse({
                    "jsonrpc": "2.0", "id": body.get("id"),
                    "error": {"code": -32020, "message": "MCP-Protocol-Version must match the request metadata"},
                }, status_code=400)
        return await call_next(request)


class OutlookMCPServer(MCPServer):
    """Use the SDK's transport with body-first version/header consistency checks."""

    def streamable_http_app(self, **kwargs):
        application = super().streamable_http_app(**kwargs)
        application.add_middleware(ModernHeaderGuard)
        # This outer limit runs before the guard reads any body, including chunks.
        application.add_middleware(
            RequestBodyLimitMiddleware,
            max_body_size=kwargs.get("max_request_body_size", DEFAULT_MAX_REQUEST_BODY_SIZE),
        )
        return application


async def mark_tool_failures(ctx: ServerRequestContext, call_next: CallNext) -> HandlerResult:
    """Preserve typed application errors and expose them as MCP tool failures."""
    result = await call_next(ctx)
    if ctx.method == "tools/call":
        if isinstance(result, types.CallToolResult):
            if (result.structured_content or {}).get("success") is False:
                return result.model_copy(update={"is_error": True})
        elif isinstance(result, dict):
            if (result.get("structuredContent") or {}).get("success") is False:
                return {**result, "isError": True}
    return result


def local_http_security() -> TransportSecuritySettings:
    """Allow only the local endpoint and reject browser cross-origin requests."""
    return TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=["127.0.0.1:8000", "localhost:8000"],
        allowed_origins=["http://127.0.0.1:8000", "http://localhost:8000"],
    )
