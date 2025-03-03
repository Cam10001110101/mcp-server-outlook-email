# Email Processing MCP Server

This MCP server provides email processing capabilities with ChromaDB integration for semantic search and SQLite for efficient storage and retrieval.

## Overview

The server processes emails from Outlook, stores them in SQLite, and generates vector embeddings for semantic search capabilities. It follows the Model Context Protocol (MCP) specification for communication with clients.

## Key Components

1. **Email Processing**
   - Outlook integration for email retrieval
   - Date range filtering
   - Multi-mailbox support
   - Support for both Inbox and Sent Items folders

2. **Storage**
   - SQLite for primary email storage
   - ChromaDB for vector embeddings
   - JSON log files for debugging

3. **Logging System**
   - MCP-compliant notifications
   - Structured JSON logging to files
   - Log rotation for file management
   - Configurable log levels

## Architecture

The server uses a layered architecture:

1. **Interface Layer**
   - MCP protocol implementation
   - Tool definitions
   - Client communication

2. **Processing Layer**
   - Email processing logic
   - Embedding generation
   - Data validation

3. **Storage Layer**
   - SQLite database
   - ChromaDB vector store
   - Log file management

## Configuration

The server requires several environment variables:
- `CHROMA_DB_PATH`: Path to ChromaDB storage
- `SQLITE_DB_PATH`: Path to SQLite database
- `EMBEDDING_BASE_URL`: Ollama server URL
- `EMBEDDING_MODEL`: Model for embeddings

## Logging

The server implements MCP-compliant logging:
- Sends structured notifications to clients
- Writes detailed logs to JSON files
- Supports standard syslog severity levels
- Implements log rotation

See [Logging Documentation](components/logging.md) for details.

## Tools

1. `process_emails`: Process emails from a date range
2. `logging/setLevel`: Configure logging verbosity

## Security

- Local data storage only
- No external API calls (except Ollama)
- Structured logging with sensitive data filtering
- Rate-limited logging
