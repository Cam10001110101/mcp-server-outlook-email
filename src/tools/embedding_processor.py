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
            self.embeddings = OllamaEmbeddings(
                model=os.getenv("EMBEDDING_MODEL"),
                base_url=os.getenv("EMBEDDING_BASE_URL")
            )
        except Exception as e:
            raise
    
    def create_email_content(self, email: Dict[str, Any]) -> str:
        """Create a formatted string of email content for embedding."""
        return f"""
Subject: {email.get('Subject', '')}
From: {email.get('SenderName', '')} <{email.get('SenderEmailAddress', '')}>
To: {email.get('To', '')}
Date: {email.get('ReceivedTime', '')}

{email.get('Body', '')}
"""

    def validate_email_data(self, email: Dict[str, Any]) -> bool:
        """Validate email data structure and content."""
        required_fields = [
            'Subject', 'SenderName', 'SenderEmailAddress', 'To', 
            'ReceivedTime', 'Folder', 'AccountName', 'Body'
        ]
        
        # Check required fields exist and are not None
        for field in required_fields:
            if field not in email or email[field] is None:
                return False
                
        # Validate dates are in ISO format
        try:
            if email['ReceivedTime']:
                datetime.fromisoformat(email['ReceivedTime'])
        except (ValueError, TypeError):
            return False
            
        return True

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
                if not self.validate_email_data(email):
                    failed_count += 1
                    continue
                    
                # Create content for embedding
                content = self.create_email_content(email)
                
                # Create metadata dictionary
                metadata = {
                    'Subject': email.get('Subject', ''),
                    'SenderName': email.get('SenderName', ''),
                    'SenderEmailAddress': email.get('SenderEmailAddress', ''),
                    'To': email.get('To', ''),
                    'ReceivedTime': email.get('ReceivedTime', ''),
                    'Folder': email.get('Folder', ''),
                    'AccountName': email.get('AccountName', '')
                }
                
                # Validate metadata can be JSON encoded
                try:
                    json.dumps(metadata)
                except (TypeError, ValueError):
                    failed_count += 1
                    continue
                
                documents.append(content)
                metadatas.append(metadata)
                ids.append(email.get('id', str(uuid.uuid4())))
                
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
