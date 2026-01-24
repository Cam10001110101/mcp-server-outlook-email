[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/cam10001110101-mcp-server-outlook-email-badge.png)](https://mseep.ai/app/cam10001110101-mcp-server-outlook-email)

# Email Processing MCP Server

A cross-platform MCP server that processes Microsoft Outlook emails, generates vector embeddings using Ollama, and provides semantic search capabilities. Works on **Windows**, **macOS**, and **any platform via Microsoft Graph API**.

The server complies with the **Model Context Protocol (MCP) 2025-06-18 specification** and uses the official MCP SDK.

## Features

### Core Capabilities
- Process emails from Outlook with date range filtering
- Store emails in SQLite database with proper connection management
- Generate vector embeddings using Ollama (nomic-embed-text)
- Semantic search across email content via MongoDB vector store
- Multi-mailbox and multi-account support
- Support for Inbox, Sent Items, and optionally Deleted Items folders

### Cross-Platform Support
- **Windows**: Native Outlook COM automation via pywin32
- **macOS**: AppleScript integration with Outlook for Mac
- **Any Platform**: Microsoft Graph API for cloud-based access (Windows, macOS, Linux, containers)

### MCP 2025-06-18 Compliance
- **Structured Tool Results**: Tools return properly typed, validated Pydantic models
- **HTTP Transport**: Supports both STDIO and HTTP (Streamable HTTP) transports
- **Protocol Negotiation**: Declares protocol version during handshake
- **Enhanced Metadata**: Tool titles and descriptions for better UI integration

### Available Tools (12+)
| Category | Tools |
|----------|-------|
| Email Processing | `process_emails` |
| Search & Analysis | `search_emails`, `analyze_email_sentiment`, `find_actionable_items` |
| Data Export | `export_email_data` (CSV, JSON, HTML, Excel) |
| Folder Management | `list_outlook_folders`, `get_folder_statistics`, `organize_emails_by_rules` |
| Contact Management | `extract_contacts` |
| Statistics | `get_email_statistics`, `check_data_consistency` |

## Prerequisites

### Required (All Platforms)
- Python 3.10 or higher
- Ollama running locally with `nomic-embed-text` model
- MongoDB server (for storing embeddings)

### Platform-Specific Requirements

| Platform | Requirement |
|----------|-------------|
| Windows | Microsoft Outlook installed + pywin32 |
| macOS | Microsoft Outlook for Mac installed |
| Graph API | Azure AD app registration with Mail.Read permission |

## Installation

### 1. Install uv (if not already installed)
```bash
pip install uv
```

### 2. Create and activate virtual environment
```bash
uv venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate
```

### 3. Install dependencies

**Core installation (required):**
```bash
uv pip install -e .
```

**Platform-specific extras:**
```bash
# Windows (adds pywin32 for COM automation)
uv pip install -e ".[windows]"

# Graph API support (cross-platform cloud access)
uv pip install -e ".[graph]"

# All optional dependencies
uv pip install -e ".[all]"
```

### 4. Install Ollama embedding model
```bash
ollama pull nomic-embed-text
```

### 5. (Graph API only) Register Azure AD Application

If using Microsoft Graph API, you need to register an application in Azure AD:

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click "New registration"
3. Name your app and select "Accounts in this organizational directory only"
4. After creation, note the **Application (client) ID** and **Directory (tenant) ID**
5. Go to "Certificates & secrets" → "New client secret" → note the secret value
6. Go to "API permissions" → "Add a permission" → "Microsoft Graph" → "Application permissions"
7. Add: `Mail.Read`, `User.Read.All` (for multi-account discovery)
8. Click "Grant admin consent"


## Configuration

Add the server to your Claude for Desktop configuration file:
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

### Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `MONGODB_URI` | MongoDB connection string | Yes |
| `SQLITE_DB_PATH` | Path to SQLite database file | Yes |
| `EMBEDDING_BASE_URL` | Ollama server URL (default: http://localhost:11434) | No |
| `EMBEDDING_MODEL` | Embedding model name (default: nomic-embed-text) | No |
| `COLLECTION_NAME` | MongoDB collection name | Yes |
| `PROCESS_DELETED_ITEMS` | Process Deleted Items folder (default: "false") | No |
| `OUTLOOK_PROVIDER` | Provider: `auto`, `windows`, `mac`, `graph` (default: "auto") | No |
| `LOCAL_TIMEZONE` | Timezone for dates (default: "UTC", e.g., "America/Chicago") | No |

**Graph API Variables (required when `OUTLOOK_PROVIDER=graph`):**

| Variable | Description |
|----------|-------------|
| `GRAPH_CLIENT_ID` | Azure AD application (client) ID |
| `GRAPH_CLIENT_SECRET` | Azure AD client secret |
| `GRAPH_TENANT_ID` | Azure AD tenant ID |
| `GRAPH_USER_EMAILS` | Mailboxes: comma-separated list or "All" for auto-discovery |

---

### Windows Configuration (COM Automation)

Uses native Outlook COM automation via pywin32.

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
        "OUTLOOK_PROVIDER": "windows",
        "LOCAL_TIMEZONE": "America/Chicago"
      }
    }
  }
}
```

### macOS Configuration (AppleScript)

Uses AppleScript to communicate with Outlook for Mac.

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

Works on any platform with Azure AD credentials. Supports single or multiple mailboxes.

**Single mailbox or specific mailboxes:**
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
        "GRAPH_USER_EMAILS": "user1@example.com,user2@example.com"
      }
    }
  }
}
```

**All mailboxes in tenant (auto-discovery):**
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
        "GRAPH_USER_EMAILS": "All"
      }
    }
  }
}
```

### Provider Auto-Detection

When `OUTLOOK_PROVIDER` is set to `auto` (default), the server automatically selects the best provider:

| Platform | Auto-Selected Provider |
|----------|----------------------|
| Windows | `windows` (COM automation) |
| macOS | `mac` (AppleScript) |
| Linux/Other | `graph` (requires Azure AD setup) |

### HTTP Transport (2025-06-18 Compliant)

For HTTP transport, run the server with the `--http` flag:

```bash
python src/mcp_server.py --http
```

This will start the server at `http://localhost:8000/mcp` with full 2025-06-18 protocol compliance including:
- Protocol version negotiation
- Structured output schemas
- HTTP header validation
- Proper error handling

The HTTP transport supports both stateful and stateless operation modes.

---

## Available Tools

### 1. process_emails
Process emails from a specified date range and return structured results:

**Input:**
```python
{
  "start_date": "2024-01-01",    # ISO format date (YYYY-MM-DD)
  "end_date": "2024-02-15",      # ISO format date (YYYY-MM-DD)
  "mailboxes": ["All"]           # List of mailbox names or ["All"] for all mailboxes
}
```

**Output (Structured):**
```python
{
  "success": true,
  "processed_count": 150,
  "retrieved_count": 200,
  "stored_count": 180,
  "failed_count": 20,
  "message": "Successfully processed 150 emails (retrieved: 200, stored: 180, failed: 20)",
  "error": null
}
```

The tool will:
1. Connect to specified Outlook mailboxes
2. Retrieve emails from Inbox and Sent Items folders (and Deleted Items if enabled)
3. Store emails in SQLite database
4. Generate embeddings using Ollama
5. Store embeddings in MongoDB for semantic search
6. Return structured results with detailed statistics


## Example Usage in Claude

```
"Process emails from February 1st to February 17th from all mailboxes"
```

## Architecture

### Provider-Based Connector Design

The server uses a **provider-based abstraction** for cross-platform email access:

```
src/connectors/
├── base.py               # OutlookConnectorBase (abstract interface)
├── factory.py            # create_connector() with auto-detection
├── windows_connector.py  # Windows COM via pywin32
├── mac_connector.py      # macOS AppleScript via osascript
└── graph_connector.py    # Microsoft Graph API (cross-platform)
```

All connectors implement the same interface, returning standardized `EmailMetadata` objects regardless of platform.

### Dual-Database Architecture

The server uses a hybrid storage approach:

**SQLite Database:**
- Primary email storage and metadata
- Full-text search capabilities
- Processing status tracking
- Date range and folder filtering
- Fast structured queries

**MongoDB:**
- Vector embeddings storage (768 dimensions)
- Semantic similarity search
- Metadata stored alongside embeddings
- Enables AI-powered search

## Error Handling

The server provides detailed error messages for common issues:
- Invalid date formats
- Connection issues with Outlook
- MongoDB errors
- Embedding generation failures with retry logic
- SQLite storage errors
- Ollama server connection issues with automatic retries

## Resource Management

The server implements proper resource management to prevent issues:
- Database connections (SQLite and MongoDB) are kept open during the server's lifetime to prevent "Cannot operate on a closed database" errors
- Connections are only closed when the server shuts down, using an atexit handler
- Destructors and context managers are used as a fallback to ensure connections are closed when objects are garbage collected
- Connection management is designed to balance resource usage with operational reliability
- Robust retry logic for external services like Ollama to handle temporary connection issues

## MCP 2025-06-18 Compliance

This server has been upgraded to comply with the MCP 2025-06-18 specification:

### Protocol Features
- **Protocol Version**: Declares `protocolVersion: "2025-06-18"` during handshake
- **Structured Output**: All tools return typed, validated Pydantic models
- **HTTP Transport**: Supports Streamable HTTP with proper header validation
- **Tool Metadata**: Enhanced tool descriptions with titles and schemas
- **Error Handling**: Structured error responses with detailed information

### Transport Support
- **STDIO**: Traditional stdin/stdout communication (default)
- **HTTP**: Streamable HTTP on localhost:8000/mcp with header validation
- **Protocol Headers**: Validates MCP-Protocol-Version and Origin headers
- **Single-Message JSON-RPC**: No batch request support per 2025-06-18 spec

### Structured Output
The `process_emails` tool returns a structured `ProcessEmailsResult` with:
- `success`: Boolean indicating operation success
- `processed_count`: Number of emails successfully processed
- `retrieved_count`: Total emails retrieved from Outlook
- `stored_count`: Number of emails stored in SQLite
- `failed_count`: Number of emails that failed processing
- `message`: Human-readable status message
- `error`: Error details if operation failed

## Security Notes

- The server only processes emails from specified mailboxes
- All data is stored locally (SQLite) and in MongoDB
- No external API calls except to local Ollama server (and Microsoft Graph if using that provider)
- Requires explicit user approval for email processing
- No sensitive email data is exposed through the MCP interface
- HTTP transport binds only to localhost for security
- Graph API uses OAuth 2.0 client credentials flow with Azure AD
- Store Azure AD credentials securely (environment variables, not in code)

## Debugging

If you encounter issues:
1. Verify emails were successfully processed (check process_emails response)
2. Ensure Ollama server is running for embedding generation
3. Check that the SQLite database is accessible
4. Verify MongoDB connection is working properly
5. Use `check_data_consistency` tool to verify SQLite/MongoDB sync

### Platform-Specific Debugging

**Windows:**
- Ensure Outlook is running and accessible
- Check that pywin32 is installed: `pip show pywin32`
- Verify COM automation works: `python -c "import win32com.client; print('OK')"`

**macOS:**
- Ensure Outlook for Mac is installed (not just the "New Outlook" web app)
- Test AppleScript access: `osascript -e 'tell application "Microsoft Outlook" to get name'`
- Grant Terminal/IDE permission in System Preferences → Security & Privacy → Automation

**Graph API:**
- Verify Azure AD app has correct permissions (Mail.Read, User.Read.All)
- Check admin consent was granted
- Test token acquisition: credentials are logged at startup
- Verify `GRAPH_USER_EMAILS` is set correctly ("All" or comma-separated emails)

## Platform Limitations

| Platform | Limitation |
|----------|------------|
| Windows | Requires Outlook desktop app running |
| macOS | "New Outlook" may have limited AppleScript support; use Graph API as fallback |
| Graph API | Requires Azure AD setup; 30-day max date range enforced |
| All | Maximum 30-day processing range per request |

## Upcoming Features

- Email summarization using LLMs
- Automatic email categorization
- Customizable email reports
- Outlook drafting email responses
- Outlook rule suggestions
- Expanded database options with Neo4j and ChromaDB integration
