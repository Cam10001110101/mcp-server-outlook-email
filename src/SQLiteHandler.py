import sqlite3
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
import json
import sys
import logging
import time
import os
from EmailMetadata import EmailMetadata

import logging

# Configure logging
logger = logging.getLogger('outlook-email.sqlite')

class SQLiteHandler:
    def __init__(self, db_path: str) -> None:
        """
        Initialize SQLite database connection and create tables if they don't exist.
        
        Args:
            db_path (str): Path to SQLite database file
        """
        try:
            logger.info(f"Initializing SQLite at {db_path}")
            self.db_path = db_path
            self.conn = self._create_connection()
            self.conn.row_factory = sqlite3.Row
            self._create_tables()
            logger.info("SQLite initialized successfully")
        except Exception as e:
            logger.error(f"Error initializing SQLite: {str(e)}", exc_info=True)
            raise

    def _create_connection(self, max_retries: int = 3) -> sqlite3.Connection:
        """Create database connection with retry logic."""
        # Ensure directory exists
        db_dir = os.path.dirname(self.db_path)
        if not os.path.exists(db_dir):
            logger.info(f"Creating directory: {db_dir}")
            os.makedirs(db_dir, exist_ok=True)
            
        for attempt in range(max_retries):
            try:
                # Use isolation_level with a value instead of None to avoid autocommit mode
                # which can cause locking issues
                return sqlite3.connect(
                    self.db_path,
                    timeout=30.0,  # 30 second timeout
                    isolation_level="IMMEDIATE"  # Use explicit transactions instead of autocommit
                )
            except Exception as e:
                if attempt == max_retries - 1:
                    raise
                logger.warning(f"Retry {attempt + 1}/{max_retries} connecting to SQLite: {str(e)}")
                time.sleep(1)

    def _create_tables(self) -> None:
        """Create necessary database tables if they don't exist."""
        cursor = self.conn.cursor()
        
        # Drop and recreate emails table
        cursor.execute('DROP TABLE IF EXISTS emails')
        
        # Create main emails table
        cursor.execute('''
        CREATE TABLE emails (
            id TEXT PRIMARY KEY,
            account TEXT NOT NULL,
            folder TEXT NOT NULL,
            subject TEXT,
            sender_name TEXT,
            sender_email TEXT,
            received_time DATETIME,
            sent_time DATETIME,
            recipients TEXT,
            is_task BOOLEAN,
            unread BOOLEAN,
            categories TEXT,
            processed BOOLEAN DEFAULT FALSE,
            last_updated DATETIME,
            body TEXT,
            attachments TEXT
        )
        ''')
        
        # Create optimized indices
        cursor.execute('CREATE INDEX idx_folder ON emails(folder)')
        cursor.execute('CREATE INDEX idx_received_time ON emails(received_time)')
        cursor.execute('CREATE INDEX idx_processed ON emails(processed)')
        
        self.conn.commit()

    def add_or_update_email(self, email: EmailMetadata, cursor: Optional[sqlite3.Cursor] = None) -> bool:
        """
        Add or update an email in the database.
        
        Args:
            email (EmailMetadata): Email metadata to store
            cursor (Optional[sqlite3.Cursor]): Optional cursor for transaction management
            
        Returns:
            bool: True if successful
        """
        try:
            # Use provided cursor or create new one
            cursor = cursor or self.conn.cursor()
            
            # Convert email to dict
            try:
                email_dict = email.to_dict()
                logger.debug(f"Processing email: {email_dict.get('Subject', 'No Subject')}")
            except Exception as e:
                logger.error(f"Error converting email to dict: {str(e)}")
                return False
            
            try:
                # Prepare data for insertion/update
                # Convert datetime objects to ISO format strings
                received_time = email_dict.get('ReceivedTime')
                sent_time = email_dict.get('SentOn')
                
                if isinstance(received_time, datetime):
                    received_time = received_time.isoformat()
                if isinstance(sent_time, datetime):
                    sent_time = sent_time.isoformat()
                
                data = {
                    'id': email_dict.get('Entry_ID'),
                    'account': email_dict.get('AccountName'),
                    'folder': email_dict.get('Folder'),
                    'subject': email_dict.get('Subject'),
                    'sender_name': email_dict.get('SenderName'),
                    'sender_email': email_dict.get('SenderEmailAddress'),
                    'received_time': received_time,
                    'sent_time': sent_time,
                    'recipients': email_dict.get('To'),
                    'is_task': bool(email_dict.get('IsMarkedAsTask')),
                    'unread': bool(email_dict.get('UnRead')),
                    'categories': email_dict.get('Categories'),
                    'processed': bool(email_dict.get('embedding')),
                    'last_updated': datetime.now().isoformat(),
                    'body': email_dict.get('Body'),
                    'attachments': email_dict.get('Attachments', '')
                }
                
                # Validate required fields
                required_fields = ['id', 'account', 'folder', 'subject', 'received_time', 'body']
                missing_fields = [field for field in required_fields if not data[field]]
                if missing_fields:
                    logger.warning(f"Missing required fields: {', '.join(missing_fields)}")
                    return False
                
            except Exception as e:
                logger.error(f"Error preparing data for SQLite: {str(e)}")
                return False
            
            # Use UPSERT syntax with retry logic
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    # Check if email exists in a transaction
                    cursor.execute('BEGIN IMMEDIATE')
                    cursor.execute('SELECT id FROM emails WHERE id = ?', (data['id'],))
                    exists = cursor.fetchone() is not None
                    
                    if exists:
                        logger.info(f"Email {data['id']} already exists, skipping")
                        cursor.execute('COMMIT')
                        return True
                    
                    # Insert new email
                    cursor.execute('''
                    INSERT INTO emails (
                        id, account, folder, subject, sender_name, sender_email,
                        received_time, sent_time, recipients, is_task, unread,
                        categories, processed, last_updated, body, attachments
                    ) VALUES (
                        :id, :account, :folder, :subject, :sender_name, :sender_email,
                        :received_time, :sent_time, :recipients, :is_task, :unread,
                        :categories, :processed, :last_updated, :body, :attachments
                    )
                    ''', data)
                    
                    cursor.execute('COMMIT')
                    logger.info(f"Successfully added email {data['id']}")
                    return True
                    
                except sqlite3.OperationalError as e:
                    cursor.execute('ROLLBACK')
                    if "database is locked" in str(e) and attempt < max_retries - 1:
                        logger.warning(f"Database locked, retry {attempt + 1}/{max_retries}")
                        time.sleep(1)
                        continue
                    logger.error(f"SQLite operational error: {str(e)}")
                    raise
                except Exception as e:
                    cursor.execute('ROLLBACK')
                    logger.error(f"Unexpected error: {str(e)}")
                    raise
                    
                except sqlite3.OperationalError as e:
                    if "database is locked" in str(e) and attempt < max_retries - 1:
                        logger.warning(f"Database locked, retry {attempt + 1}/{max_retries}")
                        time.sleep(1)
                        continue
                    raise
                    
        except Exception as e:
            logger.error(f"Error adding/updating email: {str(e)}", exc_info=True)
            self.conn.rollback()
            return False

    def get_unprocessed_emails(self, limit: int = 100) -> List[Dict[str, Any]]:
        """
        Get emails that haven't been processed (no embeddings generated).
        
        Args:
            limit (int): Maximum number of emails to return
            
        Returns:
            List[Dict]: List of unprocessed emails
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
            SELECT 
                id,
                account as AccountName,
                folder as Folder,
                subject as Subject,
                sender_name as SenderName,
                sender_email as SenderEmailAddress,
                received_time as ReceivedTime,
                sent_time as SentOn,
                recipients as "To",
                body as Body,
                COALESCE(attachments, '') as Attachments,
                is_task as IsMarkedAsTask,
                unread as UnRead,
                categories as Categories
            FROM emails 
            WHERE processed = FALSE 
            ORDER BY received_time DESC 
            LIMIT ?
            ''', (limit,))
            
            return [dict(row) for row in cursor.fetchall()]
            
        except Exception as e:
            logger.error(f"Error getting unprocessed emails: {str(e)}", exc_info=True)
            return []

    def mark_as_processed(self, email_id: str) -> bool:
        """
        Mark an email as processed after generating its embedding.
        
        Args:
            email_id (str): ID of the email to mark
            
        Returns:
            bool: True if successful
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
            UPDATE emails 
            SET processed = TRUE, 
                last_updated = ? 
            WHERE id = ?
            ''', (datetime.now().isoformat(), email_id))
            
            self.conn.commit()
            return True
            
        except Exception as e:
            logger.error(f"Error marking email as processed: {str(e)}", exc_info=True)
            self.conn.rollback()
            return False

    def get_email_by_id(self, email_id: str) -> Optional[Dict[str, Any]]:
        """
        Get a specific email by ID.
        
        Args:
            email_id (str): ID of the email to retrieve
            
        Returns:
            Optional[Dict]: Email data if found
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT * FROM emails WHERE id = ?', (email_id,))
            row = cursor.fetchone()
            return dict(row) if row else None
            
        except Exception as e:
            logger.error(f"Error getting email by ID: {str(e)}", exc_info=True)
            return None

    def get_email_count(self) -> int:
        """
        Get total number of emails in database.
        
        Returns:
            int: Number of emails
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM emails')
            return cursor.fetchone()[0]
            
        except Exception as e:
            logger.error(f"Error getting email count: {str(e)}", exc_info=True)
            return 0

    def close(self) -> None:
        """Close the database connection."""
        try:
            if hasattr(self, 'conn') and self.conn:
                self.conn.close()
                logger.info("SQLite connection closed")
        except Exception as e:
            logger.error(f"Error closing database: {str(e)}", exc_info=True)
    
    def __del__(self) -> None:
        """Destructor to ensure connection is closed when object is garbage collected."""
        self.close()
    
    def __enter__(self) -> 'SQLiteHandler':
        """Enter context manager."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Exit context manager and close connection."""
        self.close()
