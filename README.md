[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/cam10001110101-mcp-server-outlook-email-badge.png)](https://mseep.ai/app/cam10001110101-mcp-server-outlook-email)

# Email Processing MCP Server

This MCP server provides email processing capabilities with MongoDB integration for semantic search and SQLite for efficient storage and retrieval. The server has been upgraded to comply with the **Model Context Protocol (MCP) 2025-06-18 specification** and uses the official MCP SDK.

## Features

- Process emails from Outlook with date range filtering
- Store emails in SQLite database with proper connection management
- Generate vector embeddings using Ollama
- Multi-mailbox support
- Support for Inbox, Sent Items, and optionally Deleted Items folders
- **MCP 2025-06-18 Compliance**: Structured output, HTTP transport support, protocol version negotiation
- **Structured Tool Results**: Tools return properly typed, validated data structures
- **HTTP Transport**: Supports both STDIO and HTTP (Streamable HTTP) transports
- **Enhanced Metadata**: Tool titles and descriptions for better UI integration

## Upcoming Features

- Email search with semantic capabilities
- Email summarization using LLMs
- Automatic email categorization
- Customizable email reports
- Advanced filtering options
- Outlook drafting email responses
- Outlook rule suggestions
- Expanded database options with Neo4j and ChromaDB integration

## Prerequisites

- Python 3.10 or higher
- Ollama running locally (for embeddings)
- Microsoft Outlook installed
- Windows OS (for Outlook integration)
- MongoDB server (for storing embeddings)

## Installation

1. Install uv (if not already installed):
  ```bash
  pip install uv
  ```

2. Create a virtual environment:
  ```bash
  uv venv .venv
  ```

3. Activate the virtual environment:  
   
   Windows: 

    ```
    .venv\Scripts\activate
    ```

4. Install dependencies:
```bash
uv pip install -e .
```

5. The dependencies are now installed automatically from pyproject.toml, including the official MCP SDK.

6. Make sure Ollama is running locally with required models:
```bash
ollama pull nomic-embed-text
```


## Configuration

### STDIO Transport (Default)

Add the server to your Claude for Desktop configuration file:

- Windows: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "outlook-email": {
      "command": "C:/Users/username/path/to/mcp-server-outlook-email/.venv/Scripts/python",
      "args": [
        "C:/Users/username/path/to/mcp-server-outlook-email/src/mcp_server.py"
      ],
      "env": {
        "MONGODB_URI": "mongodb://localhost:27017/MCP?authSource=admin",
        "SQLITE_DB_PATH": "C:\\Users\\username\\path\\to\\mcp-server-outlook-email\\data\\emails.db",
        "EMBEDDING_BASE_URL": "http://localhost:11434",
        "EMBEDDING_MODEL": "nomic-embed-text",
        "COLLECTION_NAME": "outlook-emails",
        "PROCESS_DELETED_ITEMS": "false"
      }
    }
  }
}
```

### Tracing and Monitoring

The server has been designed to support external tracing and monitoring solutions. The MCP logging implementation has been intentionally removed in favor of a more robust tracing approach that will be implemented separately.

Note: Do not attempt to re-implement the previous logging system. A new tracing solution will be provided in the future.

Configuration fields explained:
- `command`: Full path to the Python executable in your virtual environment
- `args`: Array containing the full path to the MCP server script
- `env`: Environment variables for configuration
  - `MONGODB_URI`: MongoDB connection string
  - `SQLITE_DB_PATH`: Absolute path to SQLite database file
  - `EMBEDDING_BASE_URL`: Ollama server URL
  - `EMBEDDING_MODEL`: Model to use for embeddings
  - `LLM_MODEL`: Model to use for LLM operations
  - `COLLECTION_NAME`: Name of the MongoDB collection to use (required)
  - `PROCESS_DELETED_ITEMS`: Whether to process emails from the Deleted Items folder (optional, default: "false")
- `disabled`: Whether the server is disabled (should be false)
- `alwaysAllow`: Array of tools that don't require user confirmation
- `autoApprove`: Array of tools that can be auto-approved

Replace the paths with the actual paths on your system. Note that Windows paths in the `env` section should use double backslashes.

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

The HTTP transport supports both stateful and stateless operation modes and is the recommended approach for production deployments.

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

The server uses a hybrid search approach:
1. SQLite database for:
   - Primary email storage
   - Full-text search capabilities
   - Processing status tracking
   - Efficient filtering
   - Directory is created automatically if it doesn't exist
   - Connections are properly closed to prevent database locking

2. MongoDB for:
   - Vector embeddings storage
   - Semantic similarity search
   - Metadata filtering
   - Efficient retrieval
   - Connections are properly closed after use

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
- No external API calls except to local Ollama server
- Requires explicit user approval for email processing
- No sensitive email data is exposed through the MCP interface
- HTTP transport binds only to localhost for security

## Debugging

If you encounter issues:
1. Verify emails were successfully processed (check process_emails response)
2. Ensure Ollama server is running for embedding generation
3. Check that the SQLite database is accessible
4. Verify MongoDB connection is working properly
