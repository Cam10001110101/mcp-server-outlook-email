import logging
import time
from typing import List, Dict, Any, Optional
from pymongo import MongoClient
from pymongo.collection import Collection
from pymongo.errors import DuplicateKeyError

# Configure logging
logger = logging.getLogger('outlook-email.mongodb')

class MongoDBHandler:
    def __init__(self, connection_string: str, collection_name: str) -> None:
        """
        Initialize the MongoDBHandler with the connection string and collection name.

        Args:
            connection_string (str): MongoDB connection string
            collection_name (str): Name of the collection to manage
        """
        try:
            logger.info(f"Initializing MongoDB connection")
            self.client = MongoClient(connection_string)
            self.db = self.client.get_database()
            self.collection_name = collection_name
            self.collection = self._get_or_create_collection()
            # Create index on id field
            self.collection.create_index("id", unique=True)
            logger.info("MongoDB initialized successfully")
        except Exception as e:
            logger.error(f"Error initializing MongoDB: {str(e)}", exc_info=True)
            raise

    def _get_or_create_collection(self, max_retries: int = 3) -> Collection:
        """Get or create collection with retry logic."""
        for attempt in range(max_retries):
            try:
                return self.db[self.collection_name]
            except Exception as e:
                if attempt == max_retries - 1:
                    raise
                logger.warning(f"Retry {attempt + 1}/{max_retries} getting collection: {str(e)}")
                time.sleep(1)  # Wait before retry

    def add_embeddings(self, embeddings: List[Dict[str, Any]], job_id: Optional[str] = None) -> bool:
        """
        Add embeddings to the MongoDB collection.

        Args:
            embeddings (List[Dict]): List of embeddings to add
            job_id (str, optional): Job ID for tracking

        Returns:
            bool: True if embeddings were added successfully
        """
        try:
            # Filter out embeddings with existing IDs
            new_embeddings = []
            for embedding in embeddings:
                if not all(k in embedding for k in ['id', 'embedding', 'document', 'metadata']):
                    raise ValueError("Missing required fields in embedding")
                
                try:
                    # Check if ID already exists
                    if not self.email_exists(str(embedding['id'])):
                        # Initialize and sanitize metadata
                        embedding['metadata'] = embedding.get('metadata', {})
                        # Ensure all metadata values are primitive types
                        for key, value in embedding['metadata'].items():
                            if isinstance(value, (list, dict)):
                                embedding['metadata'][key] = str(value)
                            elif value is None:
                                embedding['metadata'][key] = ''
                        new_embeddings.append(embedding)
                except Exception as e:
                    logger.warning(f"Error checking email existence: {str(e)}")
                    continue
            
            if not new_embeddings:
                logger.info("No new embeddings to add")
                return True
            
            logger.info(f"Adding {len(new_embeddings)} embeddings to MongoDB")
            
            # Add embeddings with retry logic
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    # Convert embeddings to MongoDB documents
                    documents = []
                    for embedding in new_embeddings:
                        doc = {
                            'id': str(embedding['id']),
                            'embedding': embedding['embedding'],
                            'document': embedding['document'],
                            'metadata': embedding['metadata']
                        }
                        documents.append(doc)
                    
                    # Insert documents
                    self.collection.insert_many(documents)
                    logger.info("Successfully added embeddings to MongoDB")
                    return True
                except DuplicateKeyError:
                    logger.warning("Duplicate key found, skipping those documents")
                    return True
                except Exception as e:
                    if attempt == max_retries - 1:
                        logger.error(f"Failed to add embeddings after {max_retries} attempts: {str(e)}")
                        return False
                    logger.warning(f"Retry {attempt + 1}/{max_retries} adding embeddings: {str(e)}")
                    time.sleep(1)  # Wait before retry
            
        except Exception as e:
            logger.error(f"Error adding embeddings: {str(e)}", exc_info=True)
            return False

    def email_exists(self, entry_id: str) -> bool:
        """
        Check if an email entry exists.

        Args:
            entry_id (str): ID of the email entry

        Returns:
            bool: True if exists
        """
        try:
            result = self.collection.find_one({'id': str(entry_id)})
            return result is not None
        except Exception as e:
            logger.error(f"Error checking email existence: {str(e)}")
            return False

    def get_collection_count(self) -> int:
        """
        Get the count of documents in the collection.

        Returns:
            int: Number of documents
        """
        try:
            return self.collection.count_documents({})
        except Exception as e:
            logger.error(f"Error getting collection count: {str(e)}")
            return 0

    def get_metadata(self, entry_id: str) -> Optional[Dict[str, Any]]:
        """
        Get metadata for a specific entry.

        Args:
            entry_id (str): ID of the entry

        Returns:
            Optional[Dict[str, Any]]: The metadata if found
        """
        try:
            result = self.collection.find_one({'id': str(entry_id)})
            if result:
                return result.get('metadata')
            return None
        except Exception as e:
            logger.error(f"Error getting metadata: {str(e)}")
            return None
            
    def close(self) -> None:
        """Close the MongoDB connection."""
        try:
            if hasattr(self, 'client') and self.client:
                self.client.close()
                logger.info("MongoDB connection closed")
        except Exception as e:
            logger.error(f"Error closing MongoDB connection: {str(e)}", exc_info=True)
    
    def __del__(self) -> None:
        """Destructor to ensure connection is closed when object is garbage collected."""
        self.close()
    
    def __enter__(self) -> 'MongoDBHandler':
        """Enter context manager."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Exit context manager and close connection."""
        self.close()
