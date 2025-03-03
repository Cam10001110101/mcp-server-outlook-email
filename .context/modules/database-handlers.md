# Database Handlers

## Overview

The system integrates three specialized databases, each serving a specific purpose in the email processing and analysis pipeline.

## MongoDB Handler

### Purpose
- Store structured email metadata and vector embeddings
- Track processing status and job information
- Maintain email categories and summaries
- Store emails from Inbox, Sent Items, and optionally Deleted Items

### Implementation (MongoDBHandler.py)
```python
class MongoDBHandler:
    # Manages connection to MongoDB
    # Handles CRUD operations for email metadata
    # Tracks processing jobs and status
```

### Key Operations
- Save email metadata
- Check email existence
- Update processing status
- Retrieve email information
- Track job IDs and progress

## SQLite Handler

### Purpose
- Store primary email data
- Track processing status
- Enable efficient filtering and retrieval
- Store emails from all configured folders

### Implementation (SQLiteHandler.py)
```python
class SQLiteHandler:
    # Manages SQLite database connection
    # Handles email storage and retrieval
    # Tracks processing status
    # Supports proper connection management
```

### Key Operations
- Add embeddings
- Search similar content
- Manage collections
- Check document existence
- Retrieve embedding counts

## Neo4j Handler

### Purpose
- Store email relationship graphs
- Enable network analysis
- Track communication patterns

### Implementation (Neo4jHandler.py)
```python
class Neo4jHandler:
    # Manages Neo4j graph database
    # Creates and maintains relationship graphs
    # Provides graph querying capabilities
```

### Key Operations
- Create email relationships
- Manage vector indices
- Create and maintain constraints
- Execute graph queries
- Close connections properly

## Integration Points

### Data Flow
1. Email metadata → MongoDB
2. Vector embeddings → ChromaDB
3. Relationship data → Neo4j

### Synchronization
- Consistent Entry_IDs across databases
- Status tracking in MongoDB
- Job-based processing coordination

## Best Practices

### Connection Management
- Proper initialization and cleanup
- Connection pooling where applicable
- Error handling and retries

### Data Consistency
- Atomic operations when possible
- Transaction support where needed
- Cross-database integrity checks

### Performance
- Batch operations for efficiency
- Index optimization
- Connection reuse

## Configuration

Each handler requires specific environment variables:


### MongoDB
```
MONGODB_URI=mongodb://localhost:27017/MCP
COLLECTION_NAME=outlook-emails
```

### SQLite
```
SQLITE_DB_PATH=C:\path\to\emails.db
```

### Email Processing
```
PROCESS_DELETED_ITEMS=true|false  # Controls whether to process Deleted Items folder
```

## Upcoming Database Integrations

### Neo4j Integration
- Graph database for relationship analysis
- Communication pattern visualization
- Network analysis capabilities
- Advanced querying for complex relationships

### ChromaDB Integration
- Specialized vector database
- Optimized for embedding storage and retrieval
- Advanced similarity search capabilities
- Efficient metadata filtering

## Error Handling

- Connection failures
- Query timeouts
- Data validation errors
- Resource cleanup
- Logging and monitoring
