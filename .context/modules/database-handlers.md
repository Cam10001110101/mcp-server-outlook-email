# Database Handlers

## Overview

The system integrates three specialized databases, each serving a specific purpose in the email processing and analysis pipeline.

## MongoDB Handler

### Purpose
- Store structured email metadata
- Track processing status and job information
- Maintain email categories and summaries

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

## ChromaDB Handler

### Purpose
- Store and manage vector embeddings
- Enable semantic search capabilities
- Maintain embedding collections

### Implementation (ChromaDBHandler.py)
```python
class ChromaDBHandler:
    # Manages ChromaDB collections
    # Handles embedding storage and retrieval
    # Provides semantic search functionality
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


### ChromaDB
```
CHROMA_DB_PATH=/path/to/chromadb
CHROMA_COLLECTION_NAME=outlook-email
```

## Error Handling

- Connection failures
- Query timeouts
- Data validation errors
- Resource cleanup
- Logging and monitoring
