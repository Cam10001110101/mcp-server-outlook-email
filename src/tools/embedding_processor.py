import json
import uuid
import os
import sys
import time
import logging
from typing import List, Dict, Any, Tuple
from datetime import datetime
from langchain_ollama import OllamaEmbeddings

# Configure logging
logger = logging.getLogger('outlook-email.embedding')

class EmbeddingProcessor:
    def __init__(self, db_path: str, collection_name: str):
        """
        Initialize the embedding processor.
        
        Args:
            db_path: Path to storage
            collection_name: Name of the collection to use
        """
        # Import here to avoid circular imports
        from MongoDBHandler import MongoDBHandler
        
        # Initialize MongoDB handler
        self.mongodb_handler = MongoDBHandler(
            db_path,
            collection_name
        )
        
        # Initialize Ollama embeddings
        try:
            embedding_model = os.getenv("EMBEDDING_MODEL")
            embedding_base_url = os.getenv("EMBEDDING_BASE_URL")
            
            logger.info(f"Initializing Ollama embeddings with model={embedding_model}, base_url={embedding_base_url}")
            
            self.embeddings = OllamaEmbeddings(
                model=embedding_model,
                base_url=embedding_base_url
            )
            
            # Test connection by generating a simple embedding
            logger.info("Testing Ollama connection...")
            test_embedding = self.embeddings.embed_query("test")
            if test_embedding and len(test_embedding) > 0:
                logger.info(f"Ollama connection successful. Embedding dimension: {len(test_embedding)}")
            else:
                logger.warning("Ollama returned empty embedding for test query")
                
        except Exception as e:
            logger.error(f"Failed to initialize Ollama embeddings: {str(e)}")
            raise
    
    # Maximum characters for embedding content (nomic-embed-text has ~8192 token limit)
    # Using 4000 chars to safely fit within context limit with metadata overhead
    MAX_BODY_CHARS = 4000

    def create_email_content(self, email: Dict[str, Any]) -> str:
        """Create a formatted string of email content for embedding.

        Truncates body content to MAX_BODY_CHARS to fit within model context limits.
        """
        # Handle None values gracefully
        subject = email.get('Subject') or '(No Subject)'
        sender_name = email.get('SenderName') or 'Unknown'
        sender_email = email.get('SenderEmailAddress') or ''
        to = email.get('To') or ''
        received_time = email.get('ReceivedTime') or ''
        body = email.get('Body') or ''

        # Truncate body if too long for embedding model context
        if len(body) > self.MAX_BODY_CHARS:
            body = body[:self.MAX_BODY_CHARS] + "... [truncated]"
            logger.debug(f"Truncated email body from {len(email.get('Body', ''))} to {self.MAX_BODY_CHARS} chars")

        return f"""
Subject: {subject}
From: {sender_name} <{sender_email}>
To: {to}
Date: {received_time}

{body}
"""

    def validate_email_data(self, email: Dict[str, Any]) -> Tuple[bool, str]:
        """Validate email data structure and content.
        
        Returns:
            Tuple[bool, str]: (is_valid, reason_if_invalid)
        """
        # Only truly required fields - must exist and not be None
        required_fields = ['id', 'ReceivedTime', 'Folder', 'AccountName']
        
        # Fields that should exist but can be empty/None
        optional_fields = ['Subject', 'SenderName', 'SenderEmailAddress', 'To', 'Body']
        
        # Check required fields exist and are not None
        for field in required_fields:
            if field not in email:
                return False, f"Missing required field: {field}"
            if email[field] is None:
                return False, f"Required field is None: {field}"
        
        # Check optional fields exist (but can be None or empty)
        for field in optional_fields:
            if field not in email:
                return False, f"Missing optional field: {field}"
        
        # Must have SOME content to embed (subject or body)
        subject = email.get('Subject') or ''
        body = email.get('Body') or ''
        if not subject.strip() and not body.strip():
            return False, "Email has no subject and no body - nothing to embed"
                
        # Validate dates are in ISO format
        try:
            if email['ReceivedTime']:
                # Handle datetime strings with timezone info
                received_time = email['ReceivedTime']
                if isinstance(received_time, str):
                    # Remove timezone suffix if present for fromisoformat compatibility
                    if '+' in received_time:
                        received_time = received_time.split('+')[0]
                    datetime.fromisoformat(received_time)
        except (ValueError, TypeError) as e:
            return False, f"Invalid ReceivedTime format: {email.get('ReceivedTime')} - {str(e)}"
            
        return True, ""

    def process_batch(self, emails: List[Dict[str, Any]], batch_size: int = 4) -> Tuple[int, int]:
        """
        Process a batch of emails to generate embeddings with validation.
        
        Args:
            emails: List of email dictionaries to process
            batch_size: Size of batches for processing (default: 4)
            
        Returns:
            Tuple[int, int]: (number of successfully processed emails, number of failed emails)
        """
        documents = []
        metadatas = []
        ids = []
        failed_count = 0
        
        for i, email in enumerate(emails):
            try:
                # Validate email data
                is_valid, validation_reason = self.validate_email_data(email)
                if not is_valid:
                    logger.warning(f"Email validation failed for {email.get('id', 'unknown')}: {validation_reason}")
                    failed_count += 1
                    continue
                    
                # Create content for embedding
                content = self.create_email_content(email)
                
                # Get the email ID (used for both root-level storage and metadata)
                email_id = email.get('id', str(uuid.uuid4()))
                
                # Create metadata dictionary - IMPORTANT: include 'id' for vector search retrieval
                # Handle None values to avoid issues
                metadata = {
                    'id': email_id,  # Include ID in metadata for downstream retrieval
                    'Subject': email.get('Subject') or '',
                    'SenderName': email.get('SenderName') or '',
                    'SenderEmailAddress': email.get('SenderEmailAddress') or '',
                    'To': email.get('To') or '',
                    'ReceivedTime': email.get('ReceivedTime') or '',
                    'Folder': email.get('Folder') or '',
                    'AccountName': email.get('AccountName') or ''
                }
                
                # Validate metadata can be JSON encoded
                try:
                    json.dumps(metadata)
                except (TypeError, ValueError):
                    failed_count += 1
                    continue
                
                documents.append(content)
                metadatas.append(metadata)
                ids.append(email_id)
                
            except Exception as e:
                failed_count += 1
                continue
        
        if not documents:
            return 0, failed_count
        
        # Process documents in batches
        try:
            # Add retry logic for embedding generation
            max_embed_retries = 3
            embeddings = None
            
            for embed_attempt in range(max_embed_retries):
                try:
                    logger.info(f"Generating embeddings for {len(documents)} documents (attempt {embed_attempt + 1}/{max_embed_retries})")
                    embeddings = self.embeddings.embed_documents(documents)
                    logger.info(f"Successfully generated {len(embeddings)} embeddings")
                    break
                except Exception as e:
                    logger.error(f"Error generating embeddings (attempt {embed_attempt + 1}/{max_embed_retries}): {str(e)}")
                    if embed_attempt == max_embed_retries - 1:
                        logger.error("Failed to generate embeddings after all retries")
                        return 0, len(documents) + failed_count
                    time.sleep(2)  # Wait before retry
            
            if not embeddings:
                logger.error("No embeddings generated")
                return 0, len(documents) + failed_count
                
            # Create batch of documents to add to MongoDB
            batch = [{
                'id': id_,
                'embedding': emb,
                'document': doc,
                'metadata': meta
            } for id_, emb, doc, meta in zip(ids, embeddings, documents, metadatas)]
            
            logger.info(f"Adding {len(batch)} documents to MongoDB")
            
            # Add to MongoDB with retries
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    if self.mongodb_handler.add_embeddings(batch):
                        logger.info(f"Successfully added {len(batch)} documents to MongoDB")
                        return len(batch), failed_count
                    else:
                        logger.warning(f"Failed to add documents to MongoDB (attempt {attempt + 1}/{max_retries})")
                        if attempt < max_retries - 1:
                            time.sleep(1)
                            continue
                        return 0, len(batch) + failed_count
                except Exception as e:
                    logger.error(f"Error adding documents to MongoDB (attempt {attempt + 1}/{max_retries}): {str(e)}")
                    if attempt < max_retries - 1:
                        time.sleep(1)
                        continue
                    return 0, len(batch) + failed_count
                    
        except Exception as e:
            logger.error(f"Error processing batch: {str(e)}")
            return 0, len(documents) + failed_count
            
    def close(self) -> None:
        """Close connections."""
        try:
            # Close MongoDB connection
            if hasattr(self, 'mongodb_handler'):
                self.mongodb_handler.close()
                logger.info("MongoDB connection closed from EmbeddingProcessor")
        except Exception as e:
            logger.error(f"Error closing connections: {str(e)}")
    
    def __del__(self) -> None:
        """Destructor to ensure connections are closed when object is garbage collected."""
        self.close()
