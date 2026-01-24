"""
Vector Search Handler for Local MongoDB with Python-based similarity search.

This module provides semantic search over email embeddings stored in MongoDB
using cosine similarity computed in Python. Works with local MongoDB instances.
"""

from langchain_ollama import OllamaEmbeddings
from pymongo import MongoClient
import numpy as np
import logging
from typing import List, Tuple, Optional, Dict, Any
import re

logger = logging.getLogger(__name__)


class Document:
    """Simple document class to maintain compatibility with LangChain interface."""

    def __init__(self, page_content: str, metadata: Dict[str, Any]):
        self.page_content = page_content
        self.metadata = metadata


class VectorSearchHandler:
    """Handles vector similarity search using Python-based cosine similarity."""

    def __init__(self, connection_string: str, collection_name: str,
                 embedding_model: str, embedding_base_url: str):
        """
        Initialize the vector search handler.

        Args:
            connection_string: MongoDB connection string
            collection_name: Name of the MongoDB collection
            embedding_model: Name of the Ollama embedding model
            embedding_base_url: Base URL for Ollama server
        """
        try:
            # Initialize MongoDB client
            self.client = MongoClient(connection_string)
            db_name = connection_string.split('/')[-1].split('?')[0]
            self.collection = self.client[db_name][collection_name]

            logger.info(f"Connected to MongoDB database: {db_name}, collection: {collection_name}")

            # Initialize embeddings (same as EmbeddingProcessor)
            self.embeddings = OllamaEmbeddings(
                model=embedding_model,
                base_url=embedding_base_url
            )

            logger.info(f"Initialized Ollama embeddings with model: {embedding_model}")
            logger.info("Using Python-based cosine similarity for vector search (local MongoDB compatible)")

        except Exception as e:
            logger.error(f"Error initializing VectorSearchHandler: {str(e)}", exc_info=True)
            raise

    def _cosine_similarity(self, vec1: List[float], vec2: List[float]) -> float:
        """
        Compute cosine similarity between two vectors.

        Args:
            vec1: First vector
            vec2: Second vector

        Returns:
            Cosine similarity score (0 to 1)
        """
        try:
            # Convert to numpy arrays
            v1 = np.array(vec1)
            v2 = np.array(vec2)

            # Compute cosine similarity
            dot_product = np.dot(v1, v2)
            norm1 = np.linalg.norm(v1)
            norm2 = np.linalg.norm(v2)

            if norm1 == 0 or norm2 == 0:
                return 0.0

            similarity = dot_product / (norm1 * norm2)

            # Clamp to [0, 1] range and convert to Python float
            return float(max(0.0, min(1.0, similarity)))

        except Exception as e:
            logger.error(f"Error computing cosine similarity: {str(e)}")
            return 0.0

    def similarity_search_with_filters(
        self,
        query: str,
        k: int = 20,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
        sender: Optional[str] = None,
        folder: Optional[str] = None
    ) -> List[Tuple[Document, float]]:
        """
        Perform similarity search with metadata filters using Python-based cosine similarity.

        Args:
            query: Search query string
            k: Number of results to return
            date_from: Filter by start date (ISO format)
            date_to: Filter by end date (ISO format)
            sender: Filter by sender email address
            folder: Filter by Outlook folder name

        Returns:
            List of tuples containing (Document, similarity_score)
        """
        try:
            # Generate query embedding
            logger.info(f"Generating embedding for query: '{query}'")
            query_embedding = self.embeddings.embed_query(query)

            # Build MongoDB filter for metadata
            filter_dict = {}

            if date_from:
                filter_dict['metadata.ReceivedTime'] = {'$gte': date_from}

            if date_to:
                if 'metadata.ReceivedTime' in filter_dict:
                    filter_dict['metadata.ReceivedTime']['$lte'] = date_to
                else:
                    filter_dict['metadata.ReceivedTime'] = {'$lte': date_to}

            if sender:
                # Case-insensitive regex search for sender
                filter_dict['metadata.SenderEmailAddress'] = {
                    '$regex': re.escape(sender),
                    '$options': 'i'
                }

            if folder:
                filter_dict['metadata.Folder'] = folder

            logger.info(f"Fetching documents with {len(filter_dict)} filters")

            # Fetch documents from MongoDB with filters
            cursor = self.collection.find(filter_dict)
            documents = list(cursor)

            logger.info(f"Found {len(documents)} documents matching filters")

            if not documents:
                logger.warning("No documents found matching filters")
                return []

            # Compute similarity for each document
            results = []
            for doc in documents:
                try:
                    # Extract embedding vector
                    doc_embedding = doc.get('embedding', [])
                    if not doc_embedding:
                        continue

                    # Compute cosine similarity
                    similarity = self._cosine_similarity(query_embedding, doc_embedding)

                    # Get the document ID - check both root level and metadata for compatibility
                    # Root level 'id' is the primary source, metadata 'id' is fallback for newer documents
                    doc_id = doc.get('id', '')
                    
                    # Get metadata and ensure 'id' is included
                    metadata = doc.get('metadata', {})
                    if doc_id and 'id' not in metadata:
                        # Add root-level ID to metadata for downstream use
                        metadata = dict(metadata)  # Create a copy to avoid modifying the original
                        metadata['id'] = doc_id

                    # Create Document object with enriched metadata
                    document = Document(
                        page_content=doc.get('document', ''),
                        metadata=metadata
                    )

                    results.append((document, similarity))

                except Exception as e:
                    logger.error(f"Error processing document {doc.get('id', 'unknown')}: {str(e)}")
                    continue

            # Sort by similarity (highest first)
            results.sort(key=lambda x: x[1], reverse=True)

            # Return top-k results
            top_results = results[:k]

            logger.info(f"Vector search returned {len(top_results)} results")

            return top_results

        except Exception as e:
            logger.error(f"Error performing vector search: {str(e)}", exc_info=True)
            raise

    def check_vector_index_exists(self) -> bool:
        """
        Check if vector search is available.

        For local MongoDB, we just check if the collection has documents with embeddings.

        Returns:
            True if collection has embedded documents, False otherwise
        """
        try:
            # Count documents with embeddings
            count = self.collection.count_documents({'embedding': {'$exists': True}})

            if count > 0:
                logger.info(f"Found {count} documents with embeddings - vector search available")
                return True
            else:
                logger.warning("No documents with embeddings found - vector search not available yet")
                return False

        except Exception as e:
            logger.error(f"Error checking for embeddings: {str(e)}", exc_info=True)
            return False

    def close(self):
        """Close MongoDB connection."""
        try:
            if hasattr(self, 'client'):
                self.client.close()
                logger.info("MongoDB connection closed")
        except Exception as e:
            logger.error(f"Error closing MongoDB connection: {str(e)}", exc_info=True)
