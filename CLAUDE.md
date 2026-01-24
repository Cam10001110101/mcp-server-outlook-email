# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an MCP (Model Context Protocol) server that processes Outlook emails, generates embeddings using Ollama, and provides semantic search capabilities. It's designed to be used with Claude Desktop and complies with the **MCP 2025-06-18 specification**.

**Cross-Platform Support**: The server supports Windows, macOS, and any platform via Microsoft Graph API through a provider-based connector architecture.

## Development Setup

### Prerequisites
- Python 3.10+
- Ollama running locally with `nomic-embed-text` model
- Microsoft Outlook installed (Windows/Mac) OR Azure AD app for Graph API
- MongoDB server running locally or accessible via network

### Installation

```bash
# Install uv if needed
pip install uv

# Create and activate virtual environment
uv venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate

# Install core dependencies
uv pip install -e .

# For Windows (adds pywin32)
uv pip install -e ".[windows]"

# For Graph API support (cross-platform)
uv pip install -e ".[graph]"

# For all optional dependencies
uv pip install -e ".[all]"

# Ensure Ollama model is available
ollama pull nomic-embed-text
```

### Running the Server

**STDIO Transport (default for Claude Desktop):**
```bash
python src/mcp_server.py
```

**HTTP Transport (for testing):**
```bash
python src/mcp_server.py --http
```
Server will be available at `http://localhost:8000/mcp`

### Configuration

The server is configured via environment variables (typically set in Claude Desktop config):

**Core Configuration:**
- `MONGODB_URI`: MongoDB connection string
- `SQLITE_DB_PATH`: Path to SQLite database file
- `EMBEDDING_BASE_URL`: Ollama server URL (default: http://localhost:11434)
- `EMBEDDING_MODEL`: Embedding model name (nomic-embed-text)
- `COLLECTION_NAME`: MongoDB collection name (required)
- `PROCESS_DELETED_ITEMS`: Whether to process Deleted Items folder (default: "false")

**Platform/Provider Configuration:**
- `OUTLOOK_PROVIDER`: Provider to use - `auto`, `windows`, `mac`, or `graph` (default: "auto")
- `LOCAL_TIMEZONE`: Timezone for date conversion (default: "UTC", e.g., "America/Chicago")

**Graph API Configuration (for `graph` provider):**
- `GRAPH_CLIENT_ID`: Azure AD application (client) ID
- `GRAPH_CLIENT_SECRET`: Azure AD client secret
- `GRAPH_TENANT_ID`: Azure AD tenant ID (default: "common")
- `GRAPH_USER_EMAIL`: User email for mailbox access

## Architecture

### High-Level Data Flow

1. **Email Retrieval**: Platform connector retrieves emails (Windows COM, Mac AppleScript, or Graph API)
2. **Primary Storage**: Emails stored in SQLite (`SQLiteHandler`) for fast filtering and full-text search
3. **Embedding Generation**: `EmbeddingProcessor` creates embeddings via Ollama's nomic-embed-text
4. **Vector Storage**: Embeddings stored in MongoDB (`MongoDBHandler`) for semantic search
5. **MCP Interface**: `mcp_server.py` exposes tools via FastMCP framework

### Provider-Based Connector Architecture

The application uses a **provider-based abstraction** for cross-platform support:

```
src/connectors/
├── __init__.py           # Package exports
├── base.py               # OutlookConnectorBase ABC
├── mailbox_info.py       # Platform-agnostic MailboxInfo dataclass
├── factory.py            # create_connector() with auto-detection
├── windows_connector.py  # Windows COM via pywin32
├── mac_connector.py      # macOS AppleScript via osascript
└── graph_connector.py    # Microsoft Graph API (cross-platform)
```

**Provider Selection:**
- `auto` (default): Detects best available provider
  - Windows → `windows` (COM automation)
  - macOS → `mac` (AppleScript)
  - Other/fallback → `graph` (requires Azure AD setup)
- `windows`: Forces Windows COM (requires pywin32)
- `mac`: Forces macOS AppleScript (requires Outlook for Mac)
- `graph`: Forces Microsoft Graph API (works anywhere with Azure AD credentials)

### Hybrid Search Strategy

The application uses a **dual-database architecture**:

**SQLite Database (`SQLiteHandler`)**:
- Primary storage for all email metadata and content
- Full-text search on subject/body/sender
- Date range filtering, folder filtering, sender filtering
- Tracks processing status (whether embeddings have been generated)
- Fast SQL queries for structured filtering

**MongoDB (`MongoDBHandler`)**:
- Stores vector embeddings for semantic search
- Enables similarity search across email content
- Metadata stored alongside embeddings for retrieval

This design allows combining semantic search (MongoDB) with structured filtering (SQLite).

### Key Components

**`connectors/` Package** (`src/connectors/`):
- `OutlookConnectorBase`: Abstract base class defining the connector interface
- `MailboxInfo`: Platform-agnostic dataclass for mailbox information
- `WindowsOutlookConnector`: Windows COM implementation using pywin32
- `MacOutlookConnector`: macOS AppleScript implementation using osascript
- `GraphAPIConnector`: Microsoft Graph API implementation using MSAL
- `create_connector()`: Factory function with auto-detection

**`EmailMetadata`** (`src/EmailMetadata.py`):
- Dataclass representing email structure
- Contains validation and sanitization logic for JSON encoding
- Converts datetimes to ISO format strings
- Sanitizes text to remove control characters

**`SQLiteHandler`** (`src/SQLiteHandler.py`):
- Manages SQLite database with explicit transaction control
- Uses `isolation_level="IMMEDIATE"` to prevent locking issues
- Tracks email processing status via `processed` boolean field
- Implements retry logic for database-locked errors
- Connection stays open during server lifetime (closed via atexit handler)

**`MongoDBHandler`** (`src/MongoDBHandler.py`):
- Manages MongoDB connection and embeddings collection
- Creates unique index on `id` field
- Implements retry logic for connection and insertion failures
- Prevents duplicate embeddings by checking existence before insert

**`EmbeddingProcessor`** (`src/tools/embedding_processor.py`):
- Orchestrates embedding generation using `langchain_ollama.OllamaEmbeddings`
- Formats email content for embedding (subject + from + to + body)
- Validates email data before processing
- Implements retry logic for Ollama connection failures
- Processes emails in batches (default: 4 at a time)

**`mcp_server.py`** (`src/mcp_server.py`):
- Main MCP server using FastMCP framework
- Declares protocol version "2025-06-18" during handshake
- All tools return structured Pydantic models (not plain dicts/strings)
- Supports both STDIO and HTTP transports
- Implements 15+ MCP tools, 4 resources, and 5 prompts

### MCP Tools Structure

Tools are organized into categories:

1. **Email Processing**: `process_emails` - main tool for retrieving and processing emails
2. **Search & Analysis**: `search_emails`, `analyze_email_sentiment`, `find_actionable_items`
3. **Data Export**: `export_email_data` (CSV, JSON, HTML formats)
4. **Folder Management**: `list_outlook_folders`, `get_folder_statistics`, `organize_emails_by_rules`
5. **Contact Management**: `extract_contacts`
6. **Statistics**: `get_email_statistics`

**Important**: Most analysis tools (sentiment, action items) currently use **keyword-based pattern matching**, not LLM calls. Only embeddings are generated via Ollama.

### Resource Management

**Critical Design Pattern**: Database connections are kept open during server lifetime:
- Connections opened during `EmailProcessor.__init__()`
- Closed only on server shutdown via `atexit.register(cleanup_resources)`
- This prevents "Cannot operate on a closed database" errors
- Both SQLite and MongoDB handlers implement context managers as fallback

### Transaction Handling

**SQLite**: Uses explicit `BEGIN IMMEDIATE` / `COMMIT` / `ROLLBACK` pattern to avoid locking issues. Never relies on autocommit mode.

**MongoDB**: Implements retry logic (3 attempts with 1-second delay) for transient failures.

### Error Handling Patterns

All major operations follow this pattern:
```python
max_retries = 3
for attempt in range(max_retries):
    try:
        # operation
        break
    except SpecificError as e:
        if attempt == max_retries - 1:
            logger.error(...)
            return error_result
        time.sleep(1)
```

Applied to:
- Database connections (SQLite, MongoDB)
- Embedding generation (Ollama)
- Database inserts

## File Organization

```
src/
├── mcp_server.py              # Main MCP server, tool definitions
├── EmailMetadata.py           # Email data model with validation
├── OutlookConnector.py        # Legacy Windows connector (deprecated)
├── SQLiteHandler.py           # SQLite database operations
├── MongoDBHandler.py          # MongoDB operations
├── VectorSearchHandler.py     # Vector search implementation
├── debug_utils.py             # Debug helpers (email dumping)
├── connectors/                # Cross-platform connector package
│   ├── __init__.py
│   ├── base.py                # Abstract base class
│   ├── mailbox_info.py        # Platform-agnostic mailbox dataclass
│   ├── factory.py             # Connector factory with auto-detection
│   ├── windows_connector.py   # Windows COM connector
│   ├── mac_connector.py       # macOS AppleScript connector
│   └── graph_connector.py     # Microsoft Graph API connector
└── tools/
    └── embedding_processor.py # Embedding generation via Ollama
```

## Platform-Specific Notes

### Windows
- Uses `win32com.client` for Outlook COM automation
- Requires pywin32: `pip install pywin32`
- All file paths in config must use double backslashes in JSON

### macOS
- Uses AppleScript via `osascript` subprocess
- Requires Microsoft Outlook for Mac to be installed
- The "New Outlook" for Mac may have limited AppleScript support
- Consider using Graph API if AppleScript fails

### Graph API (Any Platform)
- Requires Azure AD app registration
- Needs `Mail.Read` and `User.Read` permissions
- Supports both delegated (user) and application (daemon) authentication
- Works on Windows, macOS, Linux, and containers

## Configuration Examples

### Windows Configuration
```json
{
  "mcpServers": {
    "outlook-email": {
      "command": "C:/path/to/.venv/Scripts/python",
      "args": ["C:/path/to/src/mcp_server.py"],
      "env": {
        "MONGODB_URI": "mongodb://localhost:27017/MCP?authSource=admin",
        "SQLITE_DB_PATH": "C:\\path\\to\\data\\emails.db",
        "EMBEDDING_BASE_URL": "http://localhost:11434",
        "EMBEDDING_MODEL": "nomic-embed-text",
        "COLLECTION_NAME": "outlook-emails",
        "PROCESS_DELETED_ITEMS": "false",
        "OUTLOOK_PROVIDER": "windows",
        "LOCAL_TIMEZONE": "America/Chicago"
      }
    }
  }
}
```

### macOS Configuration
```json
{
  "mcpServers": {
    "outlook-email": {
      "command": "/path/to/.venv/bin/python",
      "args": ["/path/to/src/mcp_server.py"],
      "env": {
        "MONGODB_URI": "mongodb://localhost:27017/MCP?authSource=admin",
        "SQLITE_DB_PATH": "/path/to/data/emails.db",
        "EMBEDDING_BASE_URL": "http://localhost:11434",
        "EMBEDDING_MODEL": "nomic-embed-text",
        "COLLECTION_NAME": "outlook-emails",
        "OUTLOOK_PROVIDER": "mac",
        "LOCAL_TIMEZONE": "America/Los_Angeles"
      }
    }
  }
}
```

### Graph API Configuration (Cross-Platform)
```json
{
  "mcpServers": {
    "outlook-email": {
      "command": "python",
      "args": ["src/mcp_server.py"],
      "env": {
        "MONGODB_URI": "mongodb://localhost:27017/MCP",
        "SQLITE_DB_PATH": "/data/emails.db",
        "EMBEDDING_BASE_URL": "http://localhost:11434",
        "EMBEDDING_MODEL": "nomic-embed-text",
        "COLLECTION_NAME": "outlook-emails",
        "OUTLOOK_PROVIDER": "graph",
        "GRAPH_CLIENT_ID": "your-azure-ad-client-id",
        "GRAPH_CLIENT_SECRET": "your-client-secret",
        "GRAPH_TENANT_ID": "your-tenant-id",
        "GRAPH_USER_EMAIL": "user@example.com"
      }
    }
  }
}
```

### MCP 2025-06-18 Compliance
- Protocol version declared during handshake
- All tools return typed Pydantic models (e.g., `ProcessEmailsResult`, `SearchEmailsResult`)
- HTTP transport supports Streamable HTTP with header validation
- No batch JSON-RPC requests (per spec)
- Tool titles and enhanced metadata for UI integration

### Logging and Tracing
**Important**: The previous MCP logging implementation was intentionally removed. External tracing solutions should be used instead. Do not attempt to re-implement internal logging beyond basic Python logging to stderr.

### LLM Usage
- **Currently**: Only Ollama embeddings (nomic-embed-text) are used
- **Future**: `LLM_MODEL` config variable exists but is not implemented
- Analysis tools (sentiment, action items) use simple keyword matching, not LLMs

### Date Range Limitations
- `process_emails` tool enforces maximum 30-day range
- All dates must be in ISO format (YYYY-MM-DD)

### Database Schema Changes
**Warning**: The SQLite `_create_tables()` method currently does `DROP TABLE IF EXISTS emails` on initialization. This means all data is lost when the server restarts. This appears to be for development purposes.

## Common Patterns

### Adding a New MCP Tool

1. Define Pydantic result model at top of `mcp_server.py`:
```python
class MyToolResult(BaseModel):
    success: bool
    data: Any
    error: Optional[str] = None
```

2. Implement tool function with decorator:
```python
@mcp.tool(title="My Tool")
async def my_tool(
    param: str,
    ctx: Context = None
) -> MyToolResult:
    """Tool description."""
    try:
        # implementation
        return MyToolResult(success=True, data=result)
    except Exception as e:
        return MyToolResult(success=False, data=None, error=str(e))
```

3. Always return structured result, never raise exceptions to MCP layer

### Processing New Emails

The flow is:
1. `process_emails` tool called with date range and mailboxes
2. Connector's `get_emails_within_date_range()` retrieves emails
3. Emails stored in SQLite via `SQLiteHandler.add_or_update_email()`
4. Unprocessed emails retrieved via `get_unprocessed_emails()`
5. `EmbeddingProcessor.process_batch()` generates embeddings
6. Embeddings stored in MongoDB via `MongoDBHandler.add_embeddings()`
7. SQLite records marked as processed via `mark_as_processed()`

### Adding a New Connector

1. Create a new file in `src/connectors/` (e.g., `icloud_connector.py`)
2. Inherit from `OutlookConnectorBase`
3. Implement all abstract methods: `provider_name`, `is_available`, `get_mailboxes`, `get_mailbox`, `get_emails_within_date_range`
4. Add the provider to `factory.py`

### Extending Database Handlers

Both handlers implement context manager protocol:
```python
with SQLiteHandler(db_path) as handler:
    # operations
    pass
# connection automatically closed
```

However, in production the handlers are kept open (not used as context managers).

## Debugging

- Logs output to stderr (captured by Claude Desktop)
- Use `logging.info()`, `logging.error()` for diagnostics
- `debug_utils.dump_email_debug()` available for detailed email inspection
- Check SQLite database directly: `sqlite3 data/emails.db`
- Verify MongoDB contents: `mongo` and inspect the collection

## Known Limitations

1. Maximum 30-day processing range
2. SQLite table dropped on server restart (development mode)
3. No actual LLM analysis (only keyword-based sentiment/action detection)
4. Semantic search not fully implemented (tool exists but may fall back to SQL search)
5. Deleted Items processing requires explicit configuration
6. Mac connector: New Outlook for Mac may have limited AppleScript support
7. Graph API: Requires Azure AD app setup with proper permissions
