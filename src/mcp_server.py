#!/usr/bin/env python3
import os
import sys
import logging
from typing import Dict, Any, List, Optional, Union
from pydantic import BaseModel, Field
from datetime import datetime as dt
from enum import Enum

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    stream=sys.stderr
)

# Create logger for this module
logger = logging.getLogger('outlook-email.mcp_server')

# Environment variables are set by the MCP config file

from datetime import datetime
from mcp.server.fastmcp import FastMCP, Context
from MongoDBHandler import MongoDBHandler
from SQLiteHandler import SQLiteHandler
from OutlookConnector import OutlookConnector
from EmailMetadata import EmailMetadata
from langchain_ollama import OllamaEmbeddings
from debug_utils import dump_email_debug

# Initialize FastMCP server with dependencies and 2025-06-18 configuration
mcp = FastMCP(
    "outlook-email",
    dependencies=[
        "pymongo",
        "langchain",
        "langchain_ollama",
        "pywin32"
    ]
)

# Configure server for 2025-06-18 protocol specification
# Note: Protocol version must be declared during initialization handshake
# The FastMCP framework will handle the protocol version negotiation

class ProcessEmailsResult(BaseModel):
    """Structured result for process_emails tool"""
    success: bool = Field(description="Whether the operation was successful")
    processed_count: int = Field(description="Number of emails processed")
    retrieved_count: int = Field(description="Number of emails retrieved")
    stored_count: int = Field(description="Number of emails stored")
    failed_count: int = Field(description="Number of emails that failed processing")
    message: str = Field(description="Human-readable status message")
    error: Optional[str] = Field(default=None, description="Error message if operation failed")

# Email data models
class EmailItem(BaseModel):
    """Individual email item with metadata"""
    id: str = Field(description="Unique email identifier")
    subject: str = Field(description="Email subject line")
    sender: str = Field(description="Email sender address")
    recipients: List[str] = Field(description="List of recipient addresses")
    sent_date: str = Field(description="Date email was sent (ISO format)")
    body_preview: str = Field(description="First 200 characters of email body")
    folder: str = Field(description="Outlook folder name")
    importance: str = Field(description="Email importance level")
    has_attachments: bool = Field(description="Whether email has attachments")
    thread_id: Optional[str] = Field(default=None, description="Conversation thread ID")

class ContactInfo(BaseModel):
    """Contact information extracted from emails"""
    email_address: str = Field(description="Contact email address")
    display_name: Optional[str] = Field(default=None, description="Contact display name")
    interaction_count: int = Field(description="Number of email interactions")
    last_interaction: str = Field(description="Date of last interaction (ISO format)")
    domains: List[str] = Field(description="Email domains for this contact")

class EmailStatistics(BaseModel):
    """Email statistics and analytics"""
    total_emails: int = Field(description="Total number of emails")
    emails_by_folder: Dict[str, int] = Field(description="Email count by folder")
    emails_by_sender: Dict[str, int] = Field(description="Email count by sender")
    daily_volume: Dict[str, int] = Field(description="Daily email volume")
    average_response_time: Optional[float] = Field(default=None, description="Average response time in hours")
    top_keywords: List[str] = Field(description="Most frequent keywords")
    sentiment_distribution: Dict[str, int] = Field(description="Distribution of email sentiments")

class CategoryEnum(str, Enum):
    """Email categories"""
    WORK = "work"
    PERSONAL = "personal"
    PROMOTIONS = "promotions"
    NOTIFICATIONS = "notifications"
    IMPORTANT = "important"
    SPAM = "spam"
    MEETINGS = "meetings"
    TRAVEL = "travel"
    FINANCE = "finance"
    OTHER = "other"

class SentimentEnum(str, Enum):
    """Email sentiment types"""
    POSITIVE = "positive"
    NEGATIVE = "negative"
    NEUTRAL = "neutral"
    URGENT = "urgent"
    FORMAL = "formal"
    CASUAL = "casual"

class SearchEmailsResult(BaseModel):
    """Structured result for search_emails tool"""
    success: bool = Field(description="Whether the search was successful")
    total_results: int = Field(description="Total number of matching emails")
    emails: List[EmailItem] = Field(description="List of matching emails")
    search_time_ms: float = Field(description="Search execution time in milliseconds")
    query_used: str = Field(description="Actual query used for search")
    error: Optional[str] = Field(default=None, description="Error message if search failed")

class EmailCategorizationResult(BaseModel):
    """Structured result for categorize_emails tool"""
    success: bool = Field(description="Whether categorization was successful")
    processed_count: int = Field(description="Number of emails processed")
    categories_assigned: Dict[CategoryEnum, int] = Field(description="Count of emails per category")
    confidence_scores: Dict[str, float] = Field(description="Average confidence per category")
    error: Optional[str] = Field(default=None, description="Error message if categorization failed")

class ContactExtractionResult(BaseModel):
    """Structured result for extract_contacts tool"""
    success: bool = Field(description="Whether contact extraction was successful")
    contacts_found: int = Field(description="Number of unique contacts found")
    contacts: List[ContactInfo] = Field(description="List of extracted contacts")
    domains_discovered: List[str] = Field(description="Unique email domains found")
    error: Optional[str] = Field(default=None, description="Error message if extraction failed")

class ThreadSummary(BaseModel):
    """Email thread summary"""
    thread_id: str = Field(description="Thread identifier")
    subject: str = Field(description="Thread subject")
    participants: List[str] = Field(description="Thread participants")
    message_count: int = Field(description="Number of messages in thread")
    date_range: str = Field(description="Date range of thread")
    summary: str = Field(description="AI-generated thread summary")
    key_points: List[str] = Field(description="Key discussion points")
    action_items: List[str] = Field(description="Extracted action items")

class SentimentAnalysisResult(BaseModel):
    """Structured result for sentiment analysis"""
    success: bool = Field(description="Whether analysis was successful")
    emails_analyzed: int = Field(description="Number of emails analyzed")
    overall_sentiment: SentimentEnum = Field(description="Overall sentiment classification")
    sentiment_breakdown: Dict[SentimentEnum, int] = Field(description="Count by sentiment type")
    confidence_score: float = Field(description="Analysis confidence (0.0-1.0)")
    trends: Dict[str, str] = Field(description="Sentiment trends over time")
    error: Optional[str] = Field(default=None, description="Error message if analysis failed")

class ActionItem(BaseModel):
    """Extracted action item from emails"""
    description: str = Field(description="Action item description")
    due_date: Optional[str] = Field(default=None, description="Due date if specified")
    priority: str = Field(description="Priority level (high/medium/low)")
    assignee: Optional[str] = Field(default=None, description="Person assigned to task")
    email_id: str = Field(description="Source email ID")
    confidence: float = Field(description="Extraction confidence (0.0-1.0)")

class ActionItemsResult(BaseModel):
    """Structured result for find_actionable_items tool"""
    success: bool = Field(description="Whether extraction was successful")
    emails_processed: int = Field(description="Number of emails processed")
    action_items: List[ActionItem] = Field(description="Extracted action items")
    meetings_found: int = Field(description="Number of meeting requests found")
    deadlines_found: int = Field(description="Number of deadlines identified")
    error: Optional[str] = Field(default=None, description="Error message if extraction failed")

class ExportFormat(str, Enum):
    """Export format options"""
    CSV = "csv"
    JSON = "json"
    PDF = "pdf"
    EXCEL = "excel"
    HTML = "html"

class ExportResult(BaseModel):
    """Structured result for export_data tool"""
    success: bool = Field(description="Whether export was successful")
    format: ExportFormat = Field(description="Export format used")
    file_path: str = Field(description="Path to exported file")
    records_exported: int = Field(description="Number of records exported")
    file_size_bytes: int = Field(description="File size in bytes")
    error: Optional[str] = Field(default=None, description="Error message if export failed")

class FolderInfo(BaseModel):
    """Outlook folder information"""
    name: str = Field(description="Folder name")
    path: str = Field(description="Full folder path")
    item_count: int = Field(description="Number of items in folder")
    unread_count: int = Field(description="Number of unread items")
    subfolder_count: int = Field(description="Number of subfolders")
    last_modified: str = Field(description="Last modification date")

class FolderManagementResult(BaseModel):
    """Structured result for folder management operations"""
    success: bool = Field(description="Whether operation was successful")
    operation: str = Field(description="Operation performed")
    folders_affected: List[FolderInfo] = Field(description="Folders that were affected")
    items_moved: int = Field(description="Number of items moved (if applicable)")
    error: Optional[str] = Field(default=None, description="Error message if operation failed")

class DraftEmailResult(BaseModel):
    """Structured result for draft_response tool"""
    success: bool = Field(description="Whether draft generation was successful")
    subject: str = Field(description="Generated email subject")
    body: str = Field(description="Generated email body")
    tone: str = Field(description="Email tone used")
    confidence: float = Field(description="Generation confidence (0.0-1.0)")
    suggestions: List[str] = Field(description="Alternative phrasings or improvements")
    error: Optional[str] = Field(default=None, description="Error message if generation failed")

class ConsistencyReportResult(BaseModel):
    """Structured result for consistency check tool"""
    success: bool = Field(description="Whether the check was successful")
    sqlite_total: int = Field(description="Total emails in SQLite")
    sqlite_processed: int = Field(description="Emails marked as processed in SQLite")
    mongodb_total: int = Field(description="Total embeddings in MongoDB")
    missing_embeddings: List[str] = Field(description="Email IDs processed but missing embeddings")
    orphaned_embeddings: List[str] = Field(description="Embedding IDs not in SQLite")
    consistent: bool = Field(description="Whether data is consistent")
    processing_rate: float = Field(description="Percentage of emails processed")
    recommendations: List[str] = Field(description="Recommended actions")
    error: Optional[str] = Field(default=None, description="Error message if check failed")

def validate_config(config: Dict[str, str]) -> None:
    """Validate required configuration values."""
    required_vars = [
        "MONGODB_URI",
        "SQLITE_DB_PATH",
        "EMBEDDING_BASE_URL",
        "EMBEDDING_MODEL",
        "COLLECTION_NAME"
    ]
    missing_vars = [var for var in required_vars if not config.get(var)]
    if missing_vars:
        raise ValueError(f"Missing required configuration: {', '.join(missing_vars)}")
    
    # Set default values for optional configuration
    if "PROCESS_DELETED_ITEMS" not in config:
        config["PROCESS_DELETED_ITEMS"] = "false"

def validate_ollama_connection(base_url: str, model: str, max_retries: int = 3) -> bool:
    """
    Validate Ollama server is running and model is available.

    Args:
        base_url: Ollama server URL
        model: Embedding model name
        max_retries: Number of connection attempts

    Returns:
        bool: True if Ollama is accessible and model exists
    """
    import requests
    import time

    for attempt in range(max_retries):
        try:
            # Check if Ollama server is running
            response = requests.get(f"{base_url}/api/tags", timeout=5)
            if response.status_code == 200:
                # Check if model is available
                models = response.json().get('models', [])
                model_names = [m.get('name', '').split(':')[0] for m in models]

                if model in model_names:
                    logger.info(f"Ollama validation successful: {model} model found")
                    return True
                else:
                    logger.warning(f"Model {model} not found. Available models: {model_names}")
                    logger.warning(f"Run: ollama pull {model}")
                    return False
            else:
                logger.warning(f"Ollama server returned status {response.status_code}")

        except requests.exceptions.ConnectionError:
            logger.warning(f"Cannot connect to Ollama at {base_url} (attempt {attempt + 1}/{max_retries})")
            if attempt < max_retries - 1:
                time.sleep(2)
                continue
        except Exception as e:
            logger.error(f"Error validating Ollama: {str(e)}")

        if attempt == max_retries - 1:
            logger.error(f"Failed to connect to Ollama after {max_retries} attempts")
            return False

    return False

def validate_mongodb_vector_index(mongodb_uri: str, collection_name: str) -> bool:
    """
    Validate MongoDB has embeddings for vector search.

    For local MongoDB with Python-based similarity search, we check if
    the collection has documents with embeddings rather than checking
    for Atlas search indexes (which don't exist in local MongoDB).

    Args:
        mongodb_uri: MongoDB connection string
        collection_name: Collection name to check

    Returns:
        bool: True if embeddings exist in collection
    """
    try:
        from pymongo import MongoClient

        client = MongoClient(mongodb_uri)
        db_name = mongodb_uri.split('/')[-1].split('?')[0]
        collection = client[db_name][collection_name]

        # Count documents with embeddings
        embedding_count = collection.count_documents({'embedding': {'$exists': True}})

        if embedding_count > 0:
            logger.info(f"Found {embedding_count} documents with embeddings - vector search ready")
            logger.info("Using Python-based cosine similarity (local MongoDB compatible)")
        else:
            logger.warning("No documents with embeddings found yet")
            logger.warning("Vector search will be available after processing emails")

        client.close()
        return embedding_count > 0

    except Exception as e:
        logger.error(f"Error checking MongoDB embeddings: {str(e)}")
        return False

class EmailProcessor:
    def __init__(self, config: Dict[str, str]):
        """
        Initialize the email processor with configuration.
        
        Args:
            config: Dictionary containing configuration values:
                - MONGODB_URI: MongoDB connection string
                - SQLITE_DB_PATH: Path to SQLite database
                - EMBEDDING_BASE_URL: Base URL for embeddings
                - EMBEDDING_MODEL: Model to use for embeddings
                - COLLECTION_NAME: Name of the MongoDB collection to use
        """
        self.config = config
        self.collection_name = config["COLLECTION_NAME"]
        
        # Initialize embedding processor
        from tools.embedding_processor import EmbeddingProcessor
        self.embedding_processor = EmbeddingProcessor(
            db_path=config["MONGODB_URI"],
            collection_name=self.collection_name
        )
        
        # Initialize SQLite handler
        self.sqlite = SQLiteHandler(config["SQLITE_DB_PATH"])

        # Initialize vector search handler
        from VectorSearchHandler import VectorSearchHandler
        try:
            self.vector_search = VectorSearchHandler(
                connection_string=config["MONGODB_URI"],
                collection_name=self.collection_name,
                embedding_model=config["EMBEDDING_MODEL"],
                embedding_base_url=config["EMBEDDING_BASE_URL"]
            )
            logger.info("Vector search handler initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize vector search: {str(e)}")
            self.vector_search = None

        # Initialize Outlook connector with deleted items setting
        process_deleted = config.get("PROCESS_DELETED_ITEMS", "false").lower() == "true"
        self.outlook = OutlookConnector(process_deleted_items=process_deleted)

    async def process_emails(
        self,
        start_date: str,
        end_date: str,
        mailboxes: List[str],
        batch_size: int = 100,
        max_to_process: Optional[int] = None,
        ctx: Context = None
    ) -> ProcessEmailsResult:
        """Process emails from the specified date range and mailboxes."""
        try:
            # Convert dates
            start = datetime.fromisoformat(start_date)
            end = datetime.fromisoformat(end_date)

            # Validate date range
            if (end - start).days > 30:
                raise ValueError("Date range cannot exceed 30 days")

            # Get mailboxes
            await ctx.report_progress(0, "Initializing email processing")
            
            if "All" in mailboxes:
                outlook_mailboxes = self.outlook.get_mailboxes()
            else:
                outlook_mailboxes = []
                for mailbox_name in mailboxes:
                    mailbox = self.outlook.get_mailbox(mailbox_name)
                    if mailbox is not None:
                        outlook_mailboxes.append(mailbox)

            if not outlook_mailboxes:
                return ProcessEmailsResult(
                    success=False,
                    processed_count=0,
                    retrieved_count=0,
                    stored_count=0,
                    failed_count=0,
                    message="No valid mailboxes found",
                    error="No valid mailboxes found"
                )

            # Include Deleted Items folder if enabled
            folder_names = ["Inbox", "Sent Items"]
            if self.outlook.process_deleted_items:
                folder_names.append("Deleted Items")
                
            await ctx.report_progress(10, f"Retrieving emails from {', '.join(folder_names)}")
            
            all_emails = []
            for i, mailbox in enumerate(outlook_mailboxes):
                try:
                    emails = self.outlook.get_emails_within_date_range(
                        folder_names,
                        start.isoformat(),
                        end.isoformat(),
                        [mailbox]
                    )
                    
                    if emails:
                        all_emails.extend(emails)
                    
                    progress = 10 + (40 * (i + 1) / len(outlook_mailboxes))
                    await ctx.report_progress(progress, f"Processing mailbox {i+1}/{len(outlook_mailboxes)}")
                except Exception:
                    continue

            if not all_emails:
                return ProcessEmailsResult(
                    success=False,
                    processed_count=0,
                    retrieved_count=len(all_emails),
                    stored_count=0,
                    failed_count=0,
                    message="No emails found in any mailbox",
                    error="No emails found in any mailbox"
                )
            
            await ctx.report_progress(50, "Storing emails in SQLite")
            
            total_stored = 0
            for i, email in enumerate(all_emails):
                if self.sqlite.add_or_update_email(email):
                    total_stored += 1
                progress = 50 + (20 * (i + 1) / len(all_emails))
                await ctx.report_progress(progress, f"Storing email {i+1}/{len(all_emails)}")
            
            if total_stored == 0:
                return ProcessEmailsResult(
                    success=False,
                    processed_count=0,
                    retrieved_count=len(all_emails),
                    stored_count=0,
                    failed_count=len(all_emails),
                    message="Failed to store any emails in SQLite",
                    error="Failed to store any emails in SQLite"
                )
            
            # Get unprocessed emails with limit
            await ctx.report_progress(70, "Fetching unprocessed emails")

            # Apply limit if specified
            limit = max_to_process if max_to_process else 10000  # Default max
            unprocessed = self.sqlite.get_unprocessed_emails(limit=limit)
            email_dicts = [email for email in unprocessed]

            await ctx.report_progress(75, f"Found {len(email_dicts)} unprocessed emails")

            if not email_dicts:
                return ProcessEmailsResult(
                    success=True,
                    processed_count=0,
                    retrieved_count=len(all_emails),
                    stored_count=total_stored,
                    failed_count=0,
                    message="No new emails to process"
                )

            # Process in batches
            total_processed = 0
            total_failed = 0

            for i in range(0, len(email_dicts), batch_size):
                batch = email_dicts[i:i+batch_size]
                progress = 75 + (15 * (i + len(batch)) / len(email_dicts))
                await ctx.report_progress(
                    progress,
                    f"Processing batch {i//batch_size + 1} ({len(batch)} emails)"
                )

                # Process batch
                processed, failed = self.embedding_processor.process_batch(batch)
                total_processed += processed
                total_failed += failed

                # Mark processed emails
                for email in batch[:processed]:
                    self.sqlite.mark_as_processed(email['id'])

                # Stop if we've hit the max_to_process limit
                if max_to_process and total_processed >= max_to_process:
                    break

            await ctx.report_progress(90, "Finalizing processing")
            
            result = ProcessEmailsResult(
                success=True,
                processed_count=total_processed,
                retrieved_count=len(all_emails),
                stored_count=total_stored,
                failed_count=total_failed,
                message=(f"Successfully processed {total_processed} emails "
                        f"(retrieved: {len(all_emails)}, stored: {total_stored}, "
                        f"failed: {total_failed})")
            )
            await ctx.report_progress(100, "Processing complete")
            return result
            
        except Exception as e:
            return ProcessEmailsResult(
                success=False,
                processed_count=0,
                retrieved_count=0,
                stored_count=0,
                failed_count=0,
                message=f"Error processing emails: {str(e)}",
                error=str(e)
            )
            
        # Note: We don't close connections here because they might be needed for future operations
        # Connections will be closed by the atexit handler when the server shuts down

try:
    # Load configuration from environment
    config = {
        "MONGODB_URI": os.environ.get("MONGODB_URI"),
        "SQLITE_DB_PATH": os.environ.get("SQLITE_DB_PATH"),
        "EMBEDDING_BASE_URL": os.environ.get("EMBEDDING_BASE_URL"),
        "EMBEDDING_MODEL": os.environ.get("EMBEDDING_MODEL"),
        "COLLECTION_NAME": os.environ.get("COLLECTION_NAME"),
        "PROCESS_DELETED_ITEMS": os.environ.get("PROCESS_DELETED_ITEMS", "false")
    }

    # Log environment variables for debugging
    logging.info("Environment variables:")
    # Redact sensitive information from MongoDB URI
    mongodb_uri = os.environ.get('MONGODB_URI', '')
    if mongodb_uri:
        # Simple redaction that keeps the host but hides credentials
        redacted_uri = mongodb_uri
        if '@' in mongodb_uri:
            # Format is typically mongodb://username:password@host:port/db
            redacted_uri = 'mongodb://' + mongodb_uri.split('@', 1)[1]
        logging.info(f"MONGODB_URI: {redacted_uri}")
    logging.info(f"SQLITE_DB_PATH: {os.environ.get('SQLITE_DB_PATH')}")
    logging.info(f"EMBEDDING_BASE_URL: {os.environ.get('EMBEDDING_BASE_URL')}")
    logging.info(f"EMBEDDING_MODEL: {os.environ.get('EMBEDDING_MODEL')}")
    logging.info(f"COLLECTION_NAME: {os.environ.get('COLLECTION_NAME')}")
    logging.info(f"PROCESS_DELETED_ITEMS: {os.environ.get('PROCESS_DELETED_ITEMS', 'false')}")
    
    # Validate configuration
    validate_config(config)

    # Validate Ollama connection
    logging.info("Validating Ollama connection...")
    ollama_ok = validate_ollama_connection(
        base_url=config["EMBEDDING_BASE_URL"],
        model=config["EMBEDDING_MODEL"]
    )
    if not ollama_ok:
        logging.error("WARNING: Ollama validation failed. Embedding generation will not work.")
        logging.error(f"Ensure Ollama is running at {config['EMBEDDING_BASE_URL']}")
        logging.error(f"And model '{config['EMBEDDING_MODEL']}' is installed: ollama pull {config['EMBEDDING_MODEL']}")

    # Validate MongoDB vector index
    logging.info("Checking MongoDB vector index...")
    vector_index_ok = validate_mongodb_vector_index(
        mongodb_uri=config["MONGODB_URI"],
        collection_name=config["COLLECTION_NAME"]
    )
    if not vector_index_ok:
        logging.error("WARNING: MongoDB vector index not found. Vector search will not work.")

    # Initialize processor
    processor = EmailProcessor(config)

    # Log startup status
    if ollama_ok and vector_index_ok:
        logging.info("✓ All systems operational")
    else:
        logging.warning("⚠ Server started with warnings - some features may not work")
    
except Exception as e:
    raise

# Register cleanup handler for server shutdown
import atexit

def cleanup_resources():
    """Clean up resources when the server shuts down."""
    try:
        if 'processor' in globals():
            # Close SQLite connection
            if hasattr(processor, 'sqlite'):
                processor.sqlite.close()
                logging.info("SQLite connection closed during shutdown")
            
            # Close MongoDB connection
            if hasattr(processor, 'embedding_processor') and hasattr(processor.embedding_processor, 'mongodb_handler'):
                processor.embedding_processor.mongodb_handler.close()
                logging.info("MongoDB connection closed during shutdown")

            # Close vector search handler
            if hasattr(processor, 'vector_search') and processor.vector_search:
                processor.vector_search.close()
                logging.info("Vector search connection closed during shutdown")
    except Exception as e:
        logging.error(f"Error during cleanup: {str(e)}")

atexit.register(cleanup_resources)

@mcp.tool(title="Process Outlook Emails")
async def process_emails(
    start_date: str,
    end_date: str,
    mailboxes: List[str],
    batch_size: int = 100,
    max_to_process: Optional[int] = None,
    ctx: Context = None
) -> ProcessEmailsResult:
    """Process emails from specified date range and mailboxes.

    Args:
        start_date: Start date in ISO format (YYYY-MM-DD)
        end_date: End date in ISO format (YYYY-MM-DD)
        mailboxes: List of mailbox names or ["All"] for all mailboxes
        batch_size: Number of emails to process in each batch (default: 100)
        max_to_process: Maximum number of unprocessed emails to process (None = all)
    """
    try:
        # Validate date formats
        try:
            datetime.fromisoformat(start_date)
            datetime.fromisoformat(end_date)
        except ValueError:
            return ProcessEmailsResult(
                success=False,
                processed_count=0,
                retrieved_count=0,
                stored_count=0,
                failed_count=0,
                message="Error: Dates must be in ISO format (YYYY-MM-DD)",
                error="Invalid date format"
            )

        result = await processor.process_emails(start_date, end_date, mailboxes, batch_size, max_to_process, ctx)
        return result
    except Exception as e:
        return ProcessEmailsResult(
            success=False,
            processed_count=0,
            retrieved_count=0,
            stored_count=0,
            failed_count=0,
            message=f"Error: {str(e)}",
            error=str(e)
        )

@mcp.tool(title="Search Emails")
async def search_emails(
    query: str,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    sender: Optional[str] = None,
    folder: Optional[str] = None,
    limit: int = 20,
    use_semantic: bool = True,
    ctx: Context = None
) -> SearchEmailsResult:
    """Search emails using semantic vector search or keyword search.

    Args:
        query: Search query (natural language for semantic, keywords for SQL)
        date_from: Start date filter in ISO format (YYYY-MM-DD)
        date_to: End date filter in ISO format (YYYY-MM-DD)
        sender: Filter by sender email address
        folder: Filter by Outlook folder name
        limit: Maximum number of results to return
        use_semantic: Use vector similarity search (True) or SQL keyword search (False)
    """
    import time
    start_time = time.time()

    try:
        # Validate date formats if provided
        if date_from:
            try:
                datetime.fromisoformat(date_from)
            except ValueError:
                return SearchEmailsResult(
                    success=False,
                    total_results=0,
                    emails=[],
                    search_time_ms=0,
                    query_used=query,
                    error="Invalid date_from format. Use YYYY-MM-DD"
                )

        if date_to:
            try:
                datetime.fromisoformat(date_to)
            except ValueError:
                return SearchEmailsResult(
                    success=False,
                    total_results=0,
                    emails=[],
                    search_time_ms=0,
                    query_used=query,
                    error="Invalid date_to format. Use YYYY-MM-DD"
                )

        # Choose search method
        if use_semantic and processor.vector_search:
            try:
                # Perform vector search with metadata filters
                results_with_scores = processor.vector_search.similarity_search_with_filters(
                    query=query,
                    k=limit,
                    date_from=date_from,
                    date_to=date_to,
                    sender=sender,
                    folder=folder
                )

                # Convert LangChain documents to EmailItem objects
                emails = []
                for doc, score in results_with_scores:
                    # Extract metadata from LangChain document
                    metadata = doc.metadata

                    # Get email ID to fetch full details from SQLite
                    email_id = metadata.get('id', '')
                    email_data = processor.sqlite.get_email_by_id(email_id)

                    if email_data:
                        email = EmailItem(
                            id=email_data.get('id', ''),
                            subject=email_data.get('subject', ''),
                            sender=email_data.get('sender_email', ''),
                            recipients=email_data.get('recipients', '').split(';') if email_data.get('recipients') else [],
                            sent_date=email_data.get('sent_time', ''),
                            body_preview=email_data.get('body', '')[:200] if email_data.get('body') else '',
                            folder=email_data.get('folder', ''),
                            importance=email_data.get('importance', 'normal'),
                            has_attachments=bool(email_data.get('has_attachments', False)),
                            thread_id=email_data.get('thread_id')
                        )
                        emails.append(email)

                search_time = (time.time() - start_time) * 1000

                return SearchEmailsResult(
                    success=True,
                    total_results=len(emails),
                    emails=emails,
                    search_time_ms=search_time,
                    query_used=f"Vector search: {query}"
                )

            except Exception as e:
                logger.error(f"Vector search failed: {str(e)}, falling back to SQL search")
                # Fall back to SQL search
                use_semantic = False

        if not use_semantic:
            # Use SQL keyword search (fallback)
            filters = {}
            if date_from:
                filters['date_from'] = date_from
            if date_to:
                filters['date_to'] = date_to
            if sender:
                filters['sender'] = sender
            if folder:
                filters['folder'] = folder

            search_results = processor.sqlite.search_emails(
                query_text=query,
                filters=filters,
                limit=limit
            )

            # Convert results to EmailItem objects
            emails = []
            for result in search_results:
                email = EmailItem(
                    id=result.get('id', ''),
                    subject=result.get('subject', ''),
                    sender=result.get('sender_email', ''),
                    recipients=result.get('recipients', '').split(';') if result.get('recipients') else [],
                    sent_date=result.get('sent_time', ''),
                    body_preview=result.get('body', '')[:200] if result.get('body') else '',
                    folder=result.get('folder', ''),
                    importance=result.get('importance', 'normal'),
                    has_attachments=bool(result.get('has_attachments', False)),
                    thread_id=result.get('thread_id')
                )
                emails.append(email)

            search_time = (time.time() - start_time) * 1000

            return SearchEmailsResult(
                success=True,
                total_results=len(emails),
                emails=emails,
                search_time_ms=search_time,
                query_used=f"SQL keyword search: {query}"
            )

    except Exception as e:
        return SearchEmailsResult(
            success=False,
            total_results=0,
            emails=[],
            search_time_ms=(time.time() - start_time) * 1000,
            query_used=query,
            error=str(e)
        )

@mcp.tool(title="Get Email Statistics")
async def get_email_statistics(
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    ctx: Context = None
) -> EmailStatistics:
    """Get comprehensive email statistics and analytics.
    
    Args:
        date_from: Start date for analysis in ISO format (YYYY-MM-DD)
        date_to: End date for analysis in ISO format (YYYY-MM-DD)
    """
    try:
        # Validate date formats if provided
        if date_from:
            try:
                datetime.fromisoformat(date_from)
            except ValueError:
                raise ValueError("Invalid date_from format. Use YYYY-MM-DD")
        
        if date_to:
            try:
                datetime.fromisoformat(date_to)
            except ValueError:
                raise ValueError("Invalid date_to format. Use YYYY-MM-DD")
        
        # Get statistics from SQLite database
        stats = processor.sqlite.get_email_statistics(
            date_from=date_from,
            date_to=date_to
        )
        
        return EmailStatistics(
            total_emails=stats.get('total_emails', 0),
            emails_by_folder=stats.get('emails_by_folder', {}),
            emails_by_sender=stats.get('emails_by_sender', {}),
            daily_volume=stats.get('daily_volume', {}),
            average_response_time=stats.get('average_response_time'),
            top_keywords=stats.get('top_keywords', []),
            sentiment_distribution=stats.get('sentiment_distribution', {})
        )
        
    except Exception as e:
        # Return empty statistics with error in message
        return EmailStatistics(
            total_emails=0,
            emails_by_folder={},
            emails_by_sender={},
            daily_volume={},
            top_keywords=[],
            sentiment_distribution={"error": str(e)}
        )

@mcp.tool(title="Extract Contacts")
async def extract_contacts(
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    min_interactions: int = 1,
    ctx: Context = None
) -> ContactExtractionResult:
    """Extract and analyze contact information from emails.
    
    Args:
        date_from: Start date for extraction in ISO format (YYYY-MM-DD)
        date_to: End date for extraction in ISO format (YYYY-MM-DD)
        min_interactions: Minimum number of interactions to include contact
    """
    try:
        # Validate date formats if provided
        if date_from:
            try:
                datetime.fromisoformat(date_from)
            except ValueError:
                return ContactExtractionResult(
                    success=False,
                    contacts_found=0,
                    contacts=[],
                    domains_discovered=[],
                    error="Invalid date_from format. Use YYYY-MM-DD"
                )
        
        if date_to:
            try:
                datetime.fromisoformat(date_to)
            except ValueError:
                return ContactExtractionResult(
                    success=False,
                    contacts_found=0,
                    contacts=[],
                    domains_discovered=[],
                    error="Invalid date_to format. Use YYYY-MM-DD"
                )
        
        # Extract contacts from database
        contacts_data = processor.sqlite.extract_contacts(
            date_from=date_from,
            date_to=date_to,
            min_interactions=min_interactions
        )
        
        contacts = []
        domains = set()
        
        for contact_data in contacts_data:
            contact_email = contact_data.get('email_address', '')
            domain = contact_email.split('@')[1] if '@' in contact_email else ''
            if domain:
                domains.add(domain)
            
            contact = ContactInfo(
                email_address=contact_email,
                display_name=contact_data.get('display_name'),
                interaction_count=contact_data.get('interaction_count', 0),
                last_interaction=contact_data.get('last_interaction', ''),
                domains=[domain] if domain else []
            )
            contacts.append(contact)
        
        return ContactExtractionResult(
            success=True,
            contacts_found=len(contacts),
            contacts=contacts,
            domains_discovered=list(domains)
        )
        
    except Exception as e:
        return ContactExtractionResult(
            success=False,
            contacts_found=0,
            contacts=[],
            domains_discovered=[],
            error=str(e)
        )

@mcp.tool(title="Analyze Email Sentiment")
async def analyze_email_sentiment(
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    sender_filter: Optional[str] = None,
    folder_filter: Optional[str] = None,
    ctx: Context = None
) -> SentimentAnalysisResult:
    """Analyze sentiment of emails using AI.
    
    Args:
        date_from: Start date for analysis in ISO format (YYYY-MM-DD)
        date_to: End date for analysis in ISO format (YYYY-MM-DD)
        sender_filter: Filter by specific sender
        folder_filter: Filter by specific folder
    """
    try:
        # Validate date formats if provided
        if date_from:
            try:
                datetime.fromisoformat(date_from)
            except ValueError:
                return SentimentAnalysisResult(
                    success=False,
                    emails_analyzed=0,
                    overall_sentiment=SentimentEnum.NEUTRAL,
                    sentiment_breakdown={},
                    confidence_score=0.0,
                    trends={},
                    error="Invalid date_from format. Use YYYY-MM-DD"
                )
        
        # Get emails for analysis
        filters = {}
        if date_from:
            filters['date_from'] = date_from
        if date_to:
            filters['date_to'] = date_to
        if sender_filter:
            filters['sender'] = sender_filter
        if folder_filter:
            filters['folder'] = folder_filter
        
        emails = processor.sqlite.get_emails_for_analysis(filters)
        
        if not emails:
            return SentimentAnalysisResult(
                success=True,
                emails_analyzed=0,
                overall_sentiment=SentimentEnum.NEUTRAL,
                sentiment_breakdown={},
                confidence_score=0.0,
                trends={}
            )
        
        # Perform sentiment analysis (simplified implementation)
        sentiment_counts = {
            SentimentEnum.POSITIVE: 0,
            SentimentEnum.NEGATIVE: 0,
            SentimentEnum.NEUTRAL: 0,
            SentimentEnum.URGENT: 0,
            SentimentEnum.FORMAL: 0,
            SentimentEnum.CASUAL: 0
        }
        
        # Simple keyword-based sentiment analysis
        positive_words = ['thank', 'great', 'excellent', 'perfect', 'appreciate', 'good', 'wonderful']
        negative_words = ['issue', 'problem', 'urgent', 'concern', 'error', 'mistake', 'unfortunately']
        formal_words = ['sincerely', 'regards', 'please', 'kindly', 'respectfully']
        casual_words = ['hey', 'hi', 'thanks', 'cool', 'awesome', 'great!']
        
        for email in emails:
            body = email.get('body', '').lower()
            subject = email.get('subject', '').lower()
            text = f"{subject} {body}"
            
            # Count sentiment indicators
            pos_count = sum(1 for word in positive_words if word in text)
            neg_count = sum(1 for word in negative_words if word in text)
            formal_count = sum(1 for word in formal_words if word in text)
            casual_count = sum(1 for word in casual_words if word in text)
            
            # Classify sentiment
            if 'urgent' in text or 'asap' in text:
                sentiment_counts[SentimentEnum.URGENT] += 1
            elif pos_count > neg_count:
                sentiment_counts[SentimentEnum.POSITIVE] += 1
            elif neg_count > pos_count:
                sentiment_counts[SentimentEnum.NEGATIVE] += 1
            else:
                sentiment_counts[SentimentEnum.NEUTRAL] += 1
            
            # Classify formality
            if formal_count > casual_count:
                sentiment_counts[SentimentEnum.FORMAL] += 1
            elif casual_count > 0:
                sentiment_counts[SentimentEnum.CASUAL] += 1
        
        # Determine overall sentiment
        max_sentiment = max(sentiment_counts.items(), key=lambda x: x[1])
        overall_sentiment = max_sentiment[0]
        
        # Calculate confidence (simplified)
        total_classified = sum(sentiment_counts.values())
        confidence = max_sentiment[1] / total_classified if total_classified > 0 else 0.0
        
        return SentimentAnalysisResult(
            success=True,
            emails_analyzed=len(emails),
            overall_sentiment=overall_sentiment,
            sentiment_breakdown=sentiment_counts,
            confidence_score=confidence,
            trends={"trend": "Analysis complete"}
        )
        
    except Exception as e:
        return SentimentAnalysisResult(
            success=False,
            emails_analyzed=0,
            overall_sentiment=SentimentEnum.NEUTRAL,
            sentiment_breakdown={},
            confidence_score=0.0,
            trends={},
            error=str(e)
        )

@mcp.tool(title="Export Email Data")
async def export_email_data(
    format: ExportFormat,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    output_path: Optional[str] = None,
    ctx: Context = None
) -> ExportResult:
    """Export email data in various formats.
    
    Args:
        format: Export format (csv, json, pdf, excel, html)
        date_from: Start date for export in ISO format (YYYY-MM-DD)
        date_to: End date for export in ISO format (YYYY-MM-DD)
        output_path: Custom output file path (optional)
    """
    try:
        import os
        import json
        import csv
        from datetime import datetime
        
        # Validate date formats if provided
        if date_from:
            try:
                datetime.fromisoformat(date_from)
            except ValueError:
                return ExportResult(
                    success=False,
                    format=format,
                    file_path="",
                    records_exported=0,
                    file_size_bytes=0,
                    error="Invalid date_from format. Use YYYY-MM-DD"
                )
        
        # Get emails for export
        filters = {}
        if date_from:
            filters['date_from'] = date_from
        if date_to:
            filters['date_to'] = date_to
        
        emails = processor.sqlite.get_emails_for_export(filters)
        
        if not emails:
            return ExportResult(
                success=False,
                format=format,
                file_path="",
                records_exported=0,
                file_size_bytes=0,
                error="No emails found for export"
            )
        
        # Generate output filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if output_path:
            file_path = output_path
        else:
            export_dir = os.path.join(os.path.dirname(processor.sqlite.db_path), "exports")
            os.makedirs(export_dir, exist_ok=True)
            file_path = os.path.join(export_dir, f"emails_export_{timestamp}.{format.value}")
        
        # Export based on format
        if format == ExportFormat.JSON:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(emails, f, indent=2, ensure_ascii=False)
        
        elif format == ExportFormat.CSV:
            if emails:
                with open(file_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=emails[0].keys())
                    writer.writeheader()
                    writer.writerows(emails)
        
        elif format == ExportFormat.HTML:
            html_content = """
            <!DOCTYPE html>
            <html>
            <head>
                <title>Email Export</title>
                <style>
                    body { font-family: Arial, sans-serif; margin: 20px; }
                    .email { border: 1px solid #ccc; margin: 10px 0; padding: 15px; }
                    .header { background-color: #f5f5f5; padding: 10px; margin-bottom: 10px; }
                    .subject { font-weight: bold; font-size: 1.1em; }
                    .meta { color: #666; font-size: 0.9em; }
                </style>
            </head>
            <body>
                <h1>Email Export Report</h1>
            """
            
            for email in emails:
                html_content += f"""
                <div class="email">
                    <div class="header">
                        <div class="subject">{email.get('subject', 'No Subject')}</div>
                        <div class="meta">From: {email.get('sender', 'Unknown')} | Date: {email.get('sent_date', 'Unknown')} | Folder: {email.get('folder', 'Unknown')}</div>
                    </div>
                    <div class="body">{email.get('body', '')[:500]}{'...' if len(email.get('body', '')) > 500 else ''}</div>
                </div>
                """
            
            html_content += "</body></html>"
            
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
        
        else:
            # For other formats, default to JSON
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(emails, f, indent=2, ensure_ascii=False)
        
        # Get file size
        file_size = os.path.getsize(file_path)
        
        return ExportResult(
            success=True,
            format=format,
            file_path=file_path,
            records_exported=len(emails),
            file_size_bytes=file_size
        )
        
    except Exception as e:
        return ExportResult(
            success=False,
            format=format,
            file_path="",
            records_exported=0,
            file_size_bytes=0,
            error=str(e)
        )

@mcp.tool(title="Find Action Items")
async def find_actionable_items(
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    sender_filter: Optional[str] = None,
    ctx: Context = None
) -> ActionItemsResult:
    """Extract actionable items, tasks, and deadlines from emails.
    
    Args:
        date_from: Start date for analysis in ISO format (YYYY-MM-DD)
        date_to: End date for analysis in ISO format (YYYY-MM-DD)
        sender_filter: Filter by specific sender
    """
    try:
        # Get emails for analysis
        filters = {}
        if date_from:
            filters['date_from'] = date_from
        if date_to:
            filters['date_to'] = date_to
        if sender_filter:
            filters['sender'] = sender_filter
        
        emails = processor.sqlite.get_emails_for_analysis(filters)
        
        if not emails:
            return ActionItemsResult(
                success=True,
                emails_processed=0,
                action_items=[],
                meetings_found=0,
                deadlines_found=0
            )
        
        action_items = []
        meetings_found = 0
        deadlines_found = 0
        
        # Keywords that indicate action items
        action_keywords = [
            'please', 'need to', 'should', 'must', 'required', 'complete',
            'finish', 'submit', 'review', 'approve', 'send', 'prepare',
            'schedule', 'arrange', 'confirm', 'follow up', 'action item'
        ]
        
        # Keywords that indicate deadlines
        deadline_keywords = [
            'deadline', 'due', 'by', 'before', 'until', 'asap', 'urgent',
            'immediately', 'today', 'tomorrow', 'this week', 'next week'
        ]
        
        # Keywords that indicate meetings
        meeting_keywords = [
            'meeting', 'call', 'conference', 'discussion', 'session',
            'appointment', 'schedule', 'agenda', 'minutes'
        ]
        
        for email in emails:
            email_id = email.get('id', '')
            subject = email.get('subject', '').lower()
            body = email.get('body', '').lower()
            sender = email.get('sender', '')
            text = f"{subject} {body}"
            
            # Check for meetings
            if any(keyword in text for keyword in meeting_keywords):
                meetings_found += 1
            
            # Check for deadlines
            has_deadline = any(keyword in text for keyword in deadline_keywords)
            if has_deadline:
                deadlines_found += 1
            
            # Extract action items
            sentences = text.split('.')
            for sentence in sentences:
                sentence = sentence.strip()
                if any(keyword in sentence for keyword in action_keywords):
                    # Determine priority
                    priority = "high" if any(word in sentence for word in ['urgent', 'asap', 'critical']) else \
                              "medium" if any(word in sentence for word in ['important', 'soon']) else "low"
                    
                    # Extract due date hints
                    due_date = None
                    if has_deadline:
                        # Simple date extraction (could be enhanced)
                        import re
                        date_pattern = r'\d{4}-\d{2}-\d{2}'
                        date_match = re.search(date_pattern, sentence)
                        if date_match:
                            due_date = date_match.group()
                    
                    # Calculate confidence based on keyword presence
                    confidence = min(0.9, len([k for k in action_keywords if k in sentence]) * 0.2 + 0.3)
                    
                    action_item = ActionItem(
                        description=sentence[:200],  # Limit description length
                        due_date=due_date,
                        priority=priority,
                        assignee=None,  # Could be enhanced to extract names
                        email_id=email_id,
                        confidence=confidence
                    )
                    action_items.append(action_item)
        
        return ActionItemsResult(
            success=True,
            emails_processed=len(emails),
            action_items=action_items,
            meetings_found=meetings_found,
            deadlines_found=deadlines_found
        )
        
    except Exception as e:
        return ActionItemsResult(
            success=False,
            emails_processed=0,
            action_items=[],
            meetings_found=0,
            deadlines_found=0,
            error=str(e)
        )

@mcp.tool(title="List Outlook Folders")
async def list_outlook_folders(ctx: Context = None) -> FolderManagementResult:
    """List all available Outlook folders and their statistics.
    
    Returns folder information including item counts, unread counts, and folder hierarchy.
    This tool helps users understand their email organization and folder structure.
    """
    try:
        folders_info = []
        
        # Get folder statistics from database
        folder_stats = processor.sqlite.get_basic_statistics().get('folder_distribution', {})
        
        # Create FolderInfo objects for each folder found in the database
        for folder_name, item_count in folder_stats.items():
            # Get unread count for this folder
            cursor = processor.sqlite.conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM emails WHERE folder = ? AND unread = 1", (folder_name,))
            unread_count = cursor.fetchone()[0]
            
            # Get last modified date
            cursor.execute("SELECT MAX(received_time) FROM emails WHERE folder = ?", (folder_name,))
            last_modified = cursor.fetchone()[0] or "Unknown"
            
            folder_info = FolderInfo(
                name=folder_name,
                path=f"\\{folder_name}",  # Simple path structure
                item_count=item_count,
                unread_count=unread_count,
                subfolder_count=0,  # Basic implementation - could be enhanced
                last_modified=last_modified
            )
            folders_info.append(folder_info)
        
        return FolderManagementResult(
            success=True,
            operation="list_folders",
            folders_affected=folders_info,
            items_moved=0
        )
        
    except Exception as e:
        return FolderManagementResult(
            success=False,
            operation="list_folders",
            folders_affected=[],
            items_moved=0,
            error=str(e)
        )

@mcp.tool(title="Get Folder Statistics")
async def get_folder_statistics(
    folder_name: str,
    include_subfolders: bool = False,
    ctx: Context = None
) -> Dict[str, Any]:
    """Get detailed statistics for a specific Outlook folder.
    
    Args:
        folder_name: Name of the folder to analyze
        include_subfolders: Whether to include subfolder statistics (future enhancement)
        
    Returns detailed statistics about the specified folder including email counts,
    date ranges, sender distribution, and activity patterns.
    """
    try:
        # Get emails from the specified folder
        emails = processor.sqlite.get_emails_by_folder(folder_name)
        
        if not emails:
            return {
                'folder_name': folder_name,
                'total_emails': 0,
                'unread_count': 0,
                'date_range': {'earliest': None, 'latest': None},
                'top_senders': {},
                'activity_by_day': {},
                'error': 'No emails found in folder or folder does not exist'
            }
        
        # Calculate statistics
        total_emails = len(emails)
        unread_count = sum(1 for email in emails if email.get('UnRead', False))
        
        # Date range
        dates = [email['ReceivedTime'] for email in emails if email.get('ReceivedTime')]
        date_range = {
            'earliest': min(dates) if dates else None,
            'latest': max(dates) if dates else None
        }
        
        # Top senders
        sender_counts = {}
        for email in emails:
            sender = email.get('SenderEmailAddress', 'Unknown')
            sender_counts[sender] = sender_counts.get(sender, 0) + 1
        
        top_senders = dict(sorted(sender_counts.items(), key=lambda x: x[1], reverse=True)[:10])
        
        # Activity by day (simple implementation)
        activity_by_day = {}
        for email in emails:
            if email.get('ReceivedTime'):
                try:
                    date_str = email['ReceivedTime'][:10]  # Extract YYYY-MM-DD
                    activity_by_day[date_str] = activity_by_day.get(date_str, 0) + 1
                except:
                    continue
        
        return {
            'folder_name': folder_name,
            'total_emails': total_emails,
            'unread_count': unread_count,
            'date_range': date_range,
            'top_senders': top_senders,
            'activity_by_day': activity_by_day,
            'average_emails_per_day': total_emails / max(1, len(activity_by_day))
        }
        
    except Exception as e:
        return {
            'folder_name': folder_name,
            'total_emails': 0,
            'unread_count': 0,
            'date_range': {'earliest': None, 'latest': None},
            'top_senders': {},
            'activity_by_day': {},
            'error': str(e)
        }

@mcp.tool(title="Organize Emails by Rules")
async def organize_emails_by_rules(
    rules: Dict[str, str],
    dry_run: bool = True,
    date_range_days: int = 30,
    ctx: Context = None
) -> FolderManagementResult:
    """Organize emails based on specified rules (simulation mode only).
    
    Args:
        rules: Dictionary of rules in format {"condition": "target_folder"}
                e.g., {"sender_contains:@company.com": "Work", "subject_contains:invoice": "Finance"}
        dry_run: Whether to simulate the organization (always True for safety)
        date_range_days: Number of days to look back for emails to organize
        
    Note: This tool only simulates email organization for safety. 
    Actual email moving would require additional Outlook automation.
    """
    try:
        from datetime import datetime, timedelta
        
        # Force dry_run to True for safety
        dry_run = True
        
        # Calculate date range
        end_date = datetime.now()
        start_date = end_date - timedelta(days=date_range_days)
        
        # Get emails within date range
        filters = {
            'date_from': start_date.isoformat(),
            'date_to': end_date.isoformat()
        }
        emails = processor.sqlite.get_emails_for_export(filters)
        
        # Simulate organization based on rules
        organization_plan = {}
        emails_to_move = 0
        
        for email in emails:
            current_folder = email.get('Folder', 'Unknown')
            target_folder = None
            
            # Apply rules
            for rule, folder in rules.items():
                if ':' in rule:
                    condition_type, condition_value = rule.split(':', 1)
                    
                    if condition_type == 'sender_contains':
                        if condition_value.lower() in email.get('SenderEmailAddress', '').lower():
                            target_folder = folder
                            break
                    elif condition_type == 'subject_contains':
                        if condition_value.lower() in email.get('Subject', '').lower():
                            target_folder = folder
                            break
                    elif condition_type == 'folder_equals':
                        if condition_value.lower() == current_folder.lower():
                            target_folder = folder
                            break
            
            if target_folder and target_folder != current_folder:
                if target_folder not in organization_plan:
                    organization_plan[target_folder] = []
                organization_plan[target_folder].append({
                    'email_id': email.get('id'),
                    'subject': email.get('Subject', ''),
                    'from_folder': current_folder,
                    'sender': email.get('SenderEmailAddress', '')
                })
                emails_to_move += 1
        
        # Create folder info for the plan
        folders_affected = []
        for folder_name, email_list in organization_plan.items():
            folder_info = FolderInfo(
                name=folder_name,
                path=f"\\{folder_name}",
                item_count=len(email_list),
                unread_count=0,  # Would need to calculate from email data
                subfolder_count=0,
                last_modified=datetime.now().isoformat()
            )
            folders_affected.append(folder_info)
        
        operation = "organize_emails_simulation" if dry_run else "organize_emails"
        
        return FolderManagementResult(
            success=True,
            operation=operation,
            folders_affected=folders_affected,
            items_moved=emails_to_move,
            error=f"Simulation complete. Would move {emails_to_move} emails. Organization plan: {organization_plan}" if dry_run else None
        )
        
    except Exception as e:
        return FolderManagementResult(
            success=False,
            operation="organize_emails",
            folders_affected=[],
            items_moved=0,
            error=str(e)
        )

@mcp.tool(title="Check Data Consistency")
async def check_data_consistency(ctx: Context = None) -> ConsistencyReportResult:
    """
    Check consistency between SQLite and MongoDB data.

    Verifies that:
    - Emails marked as processed have embeddings in MongoDB
    - Embeddings in MongoDB correspond to emails in SQLite
    - Provides recommendations for reconciliation

    Returns:
        ConsistencyReportResult with detailed consistency metrics and recommendations
    """
    try:
        # Get SQLite consistency report
        sqlite_report = processor.sqlite.get_consistency_report()

        # Verify embeddings consistency
        consistency = processor.sqlite.verify_processed_emails_have_embeddings(
            processor.embedding_processor.mongodb_handler
        )

        # Generate recommendations
        recommendations = []
        if consistency['missing_embeddings']:
            recommendations.append(
                f"Re-process {len(consistency['missing_embeddings'])} emails to generate missing embeddings"
            )
        if consistency['orphaned_embeddings']:
            recommendations.append(
                f"Clean up {len(consistency['orphaned_embeddings'])} orphaned embeddings in MongoDB"
            )
        if consistency['consistent']:
            recommendations.append("Data is consistent - no action needed")

        return ConsistencyReportResult(
            success=True,
            sqlite_total=sqlite_report['total_emails'],
            sqlite_processed=sqlite_report['processed_emails'],
            mongodb_total=consistency['embedding_count'],
            missing_embeddings=consistency['missing_embeddings'][:10],  # Limit to 10
            orphaned_embeddings=consistency['orphaned_embeddings'][:10],
            consistent=consistency['consistent'],
            processing_rate=sqlite_report['processing_rate'],
            recommendations=recommendations
        )

    except Exception as e:
        return ConsistencyReportResult(
            success=False,
            sqlite_total=0,
            sqlite_processed=0,
            mongodb_total=0,
            missing_embeddings=[],
            orphaned_embeddings=[],
            consistent=False,
            processing_rate=0.0,
            recommendations=[],
            error=str(e)
        )

# Resources for dynamic data access
@mcp.resource("email://recent/{days}", title="Recent Emails")
async def get_recent_emails(days: str) -> str:
    """Get emails from the last N days.
    
    Args:
        days: Number of days to look back
    """
    try:
        days_int = int(days)
        if days_int <= 0 or days_int > 365:
            return "Error: Days must be between 1 and 365"
        
        # Calculate date range
        from datetime import datetime, timedelta
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days_int)
        
        # Get emails from SQLite
        emails = processor.sqlite.get_emails_in_range(
            start_date.isoformat()[:10],
            end_date.isoformat()[:10]
        )
        
        # Format as readable text
        result = f"Recent emails from the last {days} days ({len(emails)} found):\n\n"
        
        for email in emails[:50]:  # Limit to 50 for readability
            result += f"Subject: {email.get('subject', 'No Subject')}\n"
            result += f"From: {email.get('sender', 'Unknown')}\n"
            result += f"Date: {email.get('sent_date', 'Unknown')}\n"
            result += f"Folder: {email.get('folder', 'Unknown')}\n"
            if email.get('body'):
                preview = email['body'][:150].replace('\n', ' ')
                result += f"Preview: {preview}...\n"
            result += "\n---\n\n"
        
        if len(emails) > 50:
            result += f"\n... and {len(emails) - 50} more emails"
        
        return result
        
    except ValueError:
        return "Error: Days must be a valid number"
    except Exception as e:
        return f"Error retrieving recent emails: {str(e)}"

@mcp.resource("email://folder/{folder_name}", title="Folder Contents")
async def get_folder_contents(folder_name: str) -> str:
    """Get contents of a specific Outlook folder.
    
    Args:
        folder_name: Name of the Outlook folder
    """
    try:
        # Decode URL-encoded folder name
        import urllib.parse
        folder_name = urllib.parse.unquote(folder_name)
        
        # Get folder contents from SQLite
        emails = processor.sqlite.get_emails_by_folder(folder_name)
        
        result = f"Contents of folder '{folder_name}' ({len(emails)} emails):\n\n"
        
        if not emails:
            return f"No emails found in folder '{folder_name}'"
        
        # Group by date for better organization
        from collections import defaultdict
        emails_by_date = defaultdict(list)
        
        for email in emails:
            date_key = email.get('sent_date', 'Unknown')[:10]  # Just the date part
            emails_by_date[date_key].append(email)
        
        for date, day_emails in sorted(emails_by_date.items(), reverse=True):
            result += f"📅 {date} ({len(day_emails)} emails)\n"
            for email in day_emails[:10]:  # Limit per day
                result += f"  • {email.get('subject', 'No Subject')} - {email.get('sender', 'Unknown')}\n"
            if len(day_emails) > 10:
                result += f"  ... and {len(day_emails) - 10} more\n"
            result += "\n"
        
        return result
        
    except Exception as e:
        return f"Error retrieving folder contents: {str(e)}"

@mcp.resource("email://statistics/summary", title="Email Statistics Summary")
async def get_statistics_summary() -> str:
    """Get a summary of email statistics and activity."""
    try:
        # Get basic statistics
        stats = processor.sqlite.get_basic_statistics()
        
        result = "📊 Email Statistics Summary\n"
        result += "=" * 30 + "\n\n"
        
        result += f"📧 Total Emails: {stats.get('total_emails', 0)}\n"
        result += f"📁 Folders: {stats.get('total_folders', 0)}\n"
        result += f"👥 Unique Contacts: {stats.get('unique_contacts', 0)}\n"
        result += f"📎 Emails with Attachments: {stats.get('emails_with_attachments', 0)}\n\n"
        
        # Top senders
        top_senders = stats.get('top_senders', [])
        if top_senders:
            result += "🏆 Top Senders:\n"
            for i, (sender, count) in enumerate(top_senders[:5], 1):
                result += f"  {i}. {sender}: {count} emails\n"
            result += "\n"
        
        # Folder distribution
        folder_stats = stats.get('folder_distribution', {})
        if folder_stats:
            result += "📁 Email Distribution by Folder:\n"
            for folder, count in sorted(folder_stats.items(), key=lambda x: x[1], reverse=True):
                result += f"  • {folder}: {count} emails\n"
            result += "\n"
        
        # Recent activity
        recent_activity = stats.get('recent_activity', {})
        if recent_activity:
            result += "⏰ Recent Activity (Last 7 days):\n"
            for date, count in sorted(recent_activity.items(), reverse=True):
                result += f"  • {date}: {count} emails\n"
        
        return result
        
    except Exception as e:
        return f"Error retrieving statistics: {str(e)}"

@mcp.resource("email://contact/{email_address}", title="Contact Interaction History")
async def get_contact_history(email_address: str) -> str:
    """Get interaction history with a specific contact.
    
    Args:
        email_address: Email address of the contact
    """
    try:
        # Decode URL-encoded email address
        import urllib.parse
        email_address = urllib.parse.unquote(email_address)
        
        # Get interaction history
        interactions = processor.sqlite.get_contact_interactions(email_address)
        
        if not interactions:
            return f"No interactions found with {email_address}"
        
        result = f"🤝 Interaction History: {email_address}\n"
        result += "=" * 50 + "\n\n"
        
        result += f"📊 Total Interactions: {len(interactions)}\n"
        
        # Count by type
        sent_count = len([i for i in interactions if i.get('type') == 'sent'])
        received_count = len([i for i in interactions if i.get('type') == 'received'])
        result += f"📤 Sent: {sent_count}\n"
        result += f"📥 Received: {received_count}\n\n"
        
        # Recent interactions
        result += "📋 Recent Interactions:\n"
        for interaction in interactions[:20]:  # Last 20 interactions
            date = interaction.get('sent_date', 'Unknown')[:10]
            subject = interaction.get('subject', 'No Subject')
            type_icon = "📤" if interaction.get('type') == 'sent' else "📥"
            result += f"  {type_icon} {date}: {subject}\n"
        
        if len(interactions) > 20:
            result += f"\n... and {len(interactions) - 20} more interactions"
        
        return result
        
    except Exception as e:
        return f"Error retrieving contact history: {str(e)}"

# Prompts for common email workflows
@mcp.prompt(title="Email Analysis")
async def email_analysis_prompt(
    email_subject: str,
    email_sender: str,
    date_range: str = "last 30 days"
) -> str:
    """Generate a comprehensive email analysis prompt.
    
    Args:
        email_subject: Subject or topic to analyze
        email_sender: Specific sender to focus on (optional)
        date_range: Time period for analysis
    """
    return f"""Please analyze emails related to "{email_subject}" from {email_sender} over the {date_range}.

Focus your analysis on:

1. **Communication Patterns**:
   - Frequency and timing of emails
   - Response times and engagement levels
   - Communication style and tone evolution

2. **Content Analysis**:
   - Key topics and themes discussed
   - Important decisions or agreements made
   - Action items and follow-up requirements

3. **Relationship Dynamics**:
   - Level of formality in communications
   - Escalation patterns or conflict resolution
   - Collaboration indicators

4. **Business Context**:
   - Project milestones and progress updates
   - Budget or resource discussions
   - Timeline and deadline tracking

Please search through the email history and provide insights based on actual email content, including specific examples and quotes where relevant."""

@mcp.prompt(title="Professional Email Response")
async def response_draft_prompt(
    original_subject: str,
    sender_name: str,
    key_points: str,
    tone: str = "professional"
) -> str:
    """Generate a professional email response template.
    
    Args:
        original_subject: Subject of the email being responded to
        sender_name: Name of the person who sent the original email
        key_points: Main points to address in the response
        tone: Desired tone (professional, friendly, formal, urgent)
    """
    return f"""Draft a {tone} email response to {sender_name} regarding "{original_subject}".

**Key Points to Address:**
{key_points}

**Response Guidelines:**
- Use a {tone} tone appropriate for business communication
- Address all key points clearly and concisely
- Include appropriate next steps or calls to action
- Maintain a constructive and solution-oriented approach
- Ensure proper email etiquette and formatting

**Structure the response with:**
1. Appropriate greeting
2. Acknowledgment of the original message
3. Clear responses to each key point
4. Next steps or proposed actions
5. Professional closing

Please generate a complete email response that I can review and send."""

@mcp.prompt(title="Meeting Follow-up")
async def meeting_followup_prompt(
    meeting_topic: str,
    attendees: str,
    meeting_date: str,
    action_items: str = ""
) -> str:
    """Generate a meeting follow-up email template.
    
    Args:
        meeting_topic: Main topic or purpose of the meeting
        attendees: List of meeting attendees
        meeting_date: Date when the meeting occurred
        action_items: Key action items from the meeting
    """
    return f"""Create a comprehensive meeting follow-up email for the "{meeting_topic}" meeting held on {meeting_date}.

**Meeting Details:**
- Topic: {meeting_topic}
- Date: {meeting_date}
- Attendees: {attendees}

**Email Should Include:**

1. **Meeting Summary**:
   - Brief recap of the main discussion points
   - Key decisions made during the meeting
   - Important announcements or updates shared

2. **Action Items**:
{action_items if action_items else "   - [Extract action items from meeting notes or discussion]"}
   - Clearly assign ownership and deadlines
   - Prioritize items by urgency and importance

3. **Next Steps**:
   - Timeline for deliverables
   - Next meeting date/time (if applicable)
   - Communication protocols for updates

4. **Documentation**:
   - Link to meeting notes or recording (if available)
   - Relevant documents or resources shared
   - Contact information for questions

Please draft a professional follow-up email that ensures all participants have a clear understanding of outcomes and next steps."""

@mcp.prompt(title="Action Item Extraction")
async def action_extraction_prompt(
    email_content: str,
    context: str = "business email"
) -> str:
    """Generate a prompt for extracting actionable items from email content.
    
    Args:
        email_content: The email content to analyze
        context: Context of the email (business, personal, project, etc.)
    """
    return f"""Analyze the following {context} and extract all actionable items, tasks, deadlines, and follow-up requirements.

**Email Content:**
{email_content}

**Please identify and categorize:**

1. **Immediate Actions** (urgent, deadline within 1-3 days):
   - Task description
   - Assigned person (if mentioned)
   - Specific deadline or timeline
   - Priority level

2. **Short-term Tasks** (deadline within 1-2 weeks):
   - Task description
   - Dependencies or prerequisites
   - Resources needed
   - Success criteria

3. **Long-term Commitments** (deadline beyond 2 weeks):
   - Project or initiative description
   - Milestones and checkpoints
   - Stakeholders involved
   - Expected outcomes

4. **Follow-up Items**:
   - Information requests
   - Decisions needed
   - Approvals required
   - Status updates

5. **Meeting/Event Items**:
   - Scheduled meetings or calls
   - Event planning requirements
   - RSVP or attendance confirmations

For each item, provide:
- Clear, actionable description
- Suggested owner/assignee
- Recommended deadline
- Priority level (High/Medium/Low)
- Any dependencies or notes"""

@mcp.prompt(title="Email Sentiment Report")
async def sentiment_report_prompt(
    sender_or_topic: str,
    time_period: str = "last 30 days",
    analysis_focus: str = "overall communication"
) -> str:
    """Generate a prompt for email sentiment analysis reporting.
    
    Args:
        sender_or_topic: Sender email or topic to analyze
        time_period: Time period for the analysis
        analysis_focus: Focus area for sentiment analysis
    """
    return f"""Conduct a comprehensive sentiment analysis of email communications related to "{sender_or_topic}" over the {time_period}, focusing on {analysis_focus}.

**Analysis Framework:**

1. **Sentiment Classification**:
   - Positive sentiment indicators (enthusiasm, satisfaction, agreement)
   - Negative sentiment indicators (frustration, disagreement, concern)
   - Neutral/professional tone patterns
   - Urgency and stress indicators

2. **Communication Tone Analysis**:
   - Formal vs. informal language usage
   - Collaborative vs. directive communication style
   - Constructive vs. critical feedback patterns
   - Emotional undertones and subtext

3. **Relationship Dynamics**:
   - Trust and rapport indicators
   - Conflict or tension signals
   - Support and assistance patterns
   - Professional boundary maintenance

4. **Trend Analysis**:
   - Sentiment evolution over time
   - Correlation with project phases or events
   - Response time patterns as sentiment indicators
   - Escalation or de-escalation patterns

5. **Business Impact Assessment**:
   - Effect on project progress and collaboration
   - Impact on team morale and productivity
   - Client satisfaction and relationship health
   - Risk factors and early warning signs

**Deliverable:**
Provide a detailed sentiment report with:
- Executive summary of key findings
- Specific examples from email content
- Trend visualization recommendations
- Actionable insights for improving communication
- Risk mitigation suggestions where applicable"""

if __name__ == "__main__":
    # Run the server with support for both STDIO and HTTP transports
    # Protocol version 2025-06-18 requires explicit version handling
    import sys
    
    # Check if HTTP transport is requested
    transport = "stdio"  # default
    if len(sys.argv) > 1 and sys.argv[1] == "--http":
        transport = "streamable-http"
        # For HTTP transport, configure headers for 2025-06-18 spec
        mcp.settings.host = "127.0.0.1"
        mcp.settings.port = 8000
        mcp.settings.mount_path = "/mcp"
        
    mcp.run(transport=transport)
