@ -0,0 +1,72 @@
# MCP Server Upgrade Plan: 2025-06-18 Specification

**Start Date:** 2025-06-30  
**Maintainer:** [Your Name Here]  
**Target:** Upgrade MCP server to comply with Model Context Protocol Specification 2025-06-18

---

## Progress Log

- **2025-06-30:** Plan created and checklist initialized.
- **2025-06-30:** Located main server entrypoints in `src/mcp_server.py` and `dxt-extension/server/src/mcp_server.py`. No explicit protocol version or batching logic found yet; next step is to audit FastMCP and tool registration for protocol handshake/version and batching support.
- **2025-06-30:** Searched all Python sources for protocolVersion, batch, jsonrpc, MCP-Protocol-Version, origin, oauth, elicitation, outputSchema, structuredContent, resource_link, _meta, context, title. No protocol version, batching, or advanced MCP 2025-06-18 features are present in the current codebase. All batch-related code is unrelated to JSON-RPC batching (it's for email/embedding processing). Next step: review FastMCP and add explicit protocol version and feature negotiation logic.
- **2025-06-30:** Located FastMCP implementation in the `fastmcp` PyPI package and its dependencies in `.venv/Lib/site-packages/fastmcp/` and `.venv/Lib/site-packages/mcp/server/fastmcp/`. The current FastMCP version does not expose protocol version negotiation or 2025-06-18 features in the server interface. Next step: determine if FastMCP or MCP server package needs to be upgraded or patched to support MCP 2025-06-18, or if a custom initialization handler must be injected.

---

## Checklist

### 1. Base Protocol
- [ ] Update initialization handshake to declare `protocolVersion: "2025-06-18"`.
- [ ] Remove JSON-RPC batching support (ban incoming batch requests).
- [ ] Enforce single-message JSON-RPC only.

### 2. Transport (HTTP + STDIO)
- [ ] HTTP: require clients to include `MCP-Protocol-Version` header on all requests.
- [ ] HTTP: validate incoming `Origin` header against allowed origins (localhost in dev).
- [ ] HTTP: ensure server binds only to `localhost` when running locally.
- [ ] STDIO: confirm message framing unchanged.

### 3. Authorization (HTTP only)
- [ ] Act as OAuth 2.1 Resource Server (implement Protected Resource Metadata RFC 9728).
- [ ] Support Resource Indicators (RFC 8707) in token requests.
- [ ] On 401 responses, include `WWW-Authenticate` header per spec.
- [ ] Validate access token audience; forbid token passthrough.

### 4. Security Considerations
- [ ] Enforce explicit user consent before tool invocations.
- [ ] Validate token audience to prevent confused-deputy attacks.
- [ ] Ensure session IDs are secure and non-deterministic (per spec 1.5.3).

### 5. Server Features
- [ ] **Resources**: support `title` field for display names.
- [ ] **Prompts**: support `title` field.
- [ ] **Tools**:
    - [ ] Declare `outputSchema` when registering tools.
    - [ ] Return structured results in a `structuredContent` field when present.
    - [ ] Allow `resource_link` as a content type in tool responses.
- [ ] **Elicitation**:
    - [ ] Declare `elicitation` capability.
    - [ ] Implement `elicitation/create` request handling (action: accept, reject, cancel).
- [ ] Add support for `_meta` and `context` fields in `completion/complete` requests.

### 6. Schema & Code Generation
- [ ] Fetch updated `schema.ts` from the 2025-06-18 spec repository.
- [ ] Regenerate any client/server models or types from the new schema.

### 7. Documentation & Tests
- [ ] Update `README.md` and examples with new version and feature descriptions.
- [ ] Add integration tests for:
    - [ ] HTTP header requirements.
    - [ ] OAuth resource metadata endpoint.
    - [ ] Structured tool output.
    - [ ] Elicitation flows.
- [ ] Record all changes in a `CHANGELOG.md`.

---

**Instructions:**  
- As you complete each item, check it off and add a timestamped log entry above.
- For each code or config change, briefly describe the file(s) and logic affected.
- If you encounter blockers, document them in the log.