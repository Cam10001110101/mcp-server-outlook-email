#!/usr/bin/env python3
import os
import sys
import logging
from typing import Dict, Any, List, Optional

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    stream=sys.stderr
)

# Environment variables are set by the MCP config file

from datetime import datetime
from fastmcp import FastMCP, Context
from MongoDBHandler import MongoDBHandler
from SQLiteHandler import SQLiteHandler
from OutlookConnector import OutlookConnector
from EmailMetadata import EmailMetadata
from langchain_ollama import OllamaEmbeddings
from debug_utils import dump_email_debug

# Initialize FastMCP server with dependencies
mcp = FastMCP(
    "outlook-email",
    dependencies=[
        "pymongo",
        "langchain",
        "langchain_ollama",
        "pywin32"
    ]
)

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
        
        # Initialize Outlook connector with deleted items setting
        process_deleted = config.get("PROCESS_DELETED_ITEMS", "false").lower() == "true"
        self.outlook = OutlookConnector(process_deleted_items=process_deleted)

    async def process_emails(
        self,
        start_date: str,
        end_date: str,
        mailboxes: List[str],
        ctx: Context
    ) -> Dict[str, Any]:
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
                return {
                    "success": False,
                    "error": "No valid mailboxes found"
                }

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
                return {
                    "success": False,
                    "error": "No emails found in any mailbox"
                }
            
            await ctx.report_progress(50, "Storing emails in SQLite")
            
            total_stored = 0
            for i, email in enumerate(all_emails):
                if self.sqlite.add_or_update_email(email):
                    total_stored += 1
                progress = 50 + (20 * (i + 1) / len(all_emails))
                await ctx.report_progress(progress, f"Storing email {i+1}/{len(all_emails)}")
            
            if total_stored == 0:
                return {
                    "success": False,
                    "error": "Failed to store any emails in SQLite"
                }
            
            unprocessed = self.sqlite.get_unprocessed_emails()
            email_dicts = [email for email in unprocessed]
            
            await ctx.report_progress(70, "Processing embeddings")
            
            if not email_dicts:
                return {
                    "success": True,
                    "processed_count": 0,
                    "message": "No new emails to process"
                }
            
            total_processed, total_failed = self.embedding_processor.process_batch(email_dicts)
            await ctx.report_progress(90, "Finalizing processing")
            
            for email in email_dicts[:total_processed]:
                self.sqlite.mark_as_processed(email['id'])
            
            result = {
                "success": True,
                "processed_count": total_processed,
                "message": (f"Successfully processed {total_processed} emails "
                          f"(retrieved: {len(all_emails)}, stored: {total_stored}, "
                          f"failed: {total_failed})")
            }
            await ctx.report_progress(100, "Processing complete")
            return result
            
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
            
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
    
    
    processor = EmailProcessor(config)
    
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
    except Exception as e:
        logging.error(f"Error during cleanup: {str(e)}")

atexit.register(cleanup_resources)

@mcp.tool()
async def process_emails(
    start_date: str,
    end_date: str,
    mailboxes: List[str],
    ctx: Context = None
) -> str:
    """Process emails from specified date range and mailboxes.
    
    Args:
        start_date: Start date in ISO format (YYYY-MM-DD)
        end_date: End date in ISO format (YYYY-MM-DD)
        mailboxes: List of mailbox names or ["All"] for all mailboxes
    """
    try:
        # Validate date formats
        try:
            datetime.fromisoformat(start_date)
            datetime.fromisoformat(end_date)
        except ValueError:
            return "Error: Dates must be in ISO format (YYYY-MM-DD)"
            
        result = await processor.process_emails(start_date, end_date, mailboxes, ctx)
        if result["success"]:
            return result["message"]
        else:
            return f"Error processing emails: {result['error']}"
    except Exception as e:
        return f"Error: {str(e)}"

if __name__ == "__main__":
    # Run the server
    mcp.run()
