[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/cam10001110101-mcp-server-outlook-email-badge.png)](https://mseep.ai/app/cam10001110101-mcp-server-outlook-email)

# Email Processing MCP Server

This MCP server provides email processing capabilities with MongoDB integration for semantic search and SQLite for efficient storage and retrieval.

## Features

- Process emails from Outlook with date range filtering
- Store emails in SQLite database with proper connection management
- Generate vector embeddings using Ollama
- Multi-mailbox support
- Support for Inbox, Sent Items, and optionally Deleted Items folders

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

   
    macOS/Linux: 

    ```python
    source .venv/bin/activate
    ```

4. Install dependencies:
```bash
uv pip install -e .
```

5. Install the fastmcp package:
```bash
uv pip install fastmcp
```

6. Make sure Ollama is running locally with required models:
```bash
ollama pull nomic-embed-text
```


## Configuration

Add the server to your Claude for Desktop configuration file:

- Windows: `%APPDATA%\Claude\claude_desktop_config.json`
- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

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

## Available Tools

### 1. process_emails
Process emails from a specified date range:
```python
{
  "start_date": "2024-01-01",    # ISO format date (YYYY-MM-DD)
  "end_date": "2024-02-15",      # ISO format date (YYYY-MM-DD)
  "mailboxes": ["All"]           # List of mailbox names or ["All"] for all mailboxes
}
```

The tool will:
1. Connect to specified Outlook mailboxes
2. Retrieve emails from Inbox and Sent Items folders (and Deleted Items if enabled)
3. Store emails in SQLite database
4. Generate embeddings using Ollama
5. Store embeddings in MongoDB for semantic search


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

## Security Notes

- The server only processes emails from specified mailboxes
- All data is stored locally (SQLite) and in MongoDB
- No external API calls except to local Ollama server
- Requires explicit user approval for email processing
- No sensitive email data is exposed through the MCP interface

## Debugging

If you encounter issues:
1. Verify emails were successfully processed (check process_emails response)
2. Ensure Ollama server is running for embedding generation
3. Check that the SQLite database is accessible
4. Verify MongoDB connection is working properly
