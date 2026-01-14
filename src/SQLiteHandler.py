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

    def get_emails_for_export(self, filters: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
        """
        Get emails for export based on filters.
        
        Args:
            filters (Optional[Dict]): Filter criteria including:
                - date_from: Start date (ISO format)
                - date_to: End date (ISO format)
                - folder: Folder name filter
                - sender: Sender email filter
                
        Returns:
            List[Dict]: List of emails matching filters
        """
        try:
            cursor = self.conn.cursor()
            
            # Build query based on filters
            query = '''
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
            WHERE 1=1
            '''
            params = []
            
            if filters:
                if filters.get('date_from'):
                    query += ' AND received_time >= ?'
                    params.append(filters['date_from'])
                
                if filters.get('date_to'):
                    query += ' AND received_time <= ?'
                    params.append(filters['date_to'])
                
                if filters.get('folder'):
                    query += ' AND folder = ?'
                    params.append(filters['folder'])
                
                if filters.get('sender'):
                    query += ' AND sender_email LIKE ?'
                    params.append(f"%{filters['sender']}%")
            
            query += ' ORDER BY received_time DESC'
            
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]
            
        except Exception as e:
            logger.error(f"Error getting emails for export: {str(e)}", exc_info=True)
            return []

    def get_emails_for_analysis(self, filters: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
        """
        Get emails for sentiment analysis based on filters.
        
        Args:
            filters (Optional[Dict]): Filter criteria including:
                - date_from: Start date (ISO format)
                - date_to: End date (ISO format)
                - folder: Folder name filter
                - sender: Sender email filter
                
        Returns:
            List[Dict]: List of emails with body text for analysis
        """
        try:
            cursor = self.conn.cursor()
            
            query = '''
            SELECT 
                id,
                subject as Subject,
                sender_email as SenderEmailAddress,
                received_time as ReceivedTime,
                body as Body,
                folder as Folder
            FROM emails 
            WHERE body IS NOT NULL AND body != ''
            '''
            params = []
            
            if filters:
                if filters.get('date_from'):
                    query += ' AND received_time >= ?'
                    params.append(filters['date_from'])
                
                if filters.get('date_to'):
                    query += ' AND received_time <= ?'
                    params.append(filters['date_to'])
                
                if filters.get('folder'):
                    query += ' AND folder = ?'
                    params.append(filters['folder'])
                
                if filters.get('sender'):
                    query += ' AND sender_email LIKE ?'
                    params.append(f"%{filters['sender']}%")
            
            query += ' ORDER BY received_time DESC'
            
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]
            
        except Exception as e:
            logger.error(f"Error getting emails for analysis: {str(e)}", exc_info=True)
            return []

    def search_emails(self, query_text: str, filters: Optional[Dict[str, Any]] = None, limit: int = 50) -> List[Dict[str, Any]]:
        """
        Search emails by text content.
        
        Args:
            query_text (str): Text to search for
            filters (Optional[Dict]): Additional filter criteria
            limit (int): Maximum number of results
            
        Returns:
            List[Dict]: List of matching emails
        """
        try:
            cursor = self.conn.cursor()
            
            base_query = '''
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
            WHERE (subject LIKE ? OR body LIKE ? OR sender_name LIKE ?)
            '''
            
            params = [f"%{query_text}%", f"%{query_text}%", f"%{query_text}%"]
            
            if filters:
                if filters.get('date_from'):
                    base_query += ' AND received_time >= ?'
                    params.append(filters['date_from'])
                
                if filters.get('date_to'):
                    base_query += ' AND received_time <= ?'
                    params.append(filters['date_to'])
                
                if filters.get('folder'):
                    base_query += ' AND folder = ?'
                    params.append(filters['folder'])
                
                if filters.get('sender'):
                    base_query += ' AND sender_email LIKE ?'
                    params.append(f"%{filters['sender']}%")
            
            base_query += ' ORDER BY received_time DESC LIMIT ?'
            params.append(limit)
            
            cursor.execute(base_query, params)
            return [dict(row) for row in cursor.fetchall()]
            
        except Exception as e:
            logger.error(f"Error searching emails: {str(e)}", exc_info=True)
            return []

    def get_email_statistics(self, date_from: Optional[str] = None, date_to: Optional[str] = None) -> Dict[str, Any]:
        """
        Get email statistics for the specified date range.
        
        Args:
            date_from (Optional[str]): Start date (ISO format)
            date_to (Optional[str]): End date (ISO format)
            
        Returns:
            Dict: Email statistics
        """
        try:
            cursor = self.conn.cursor()
            
            # Build base query
            base_where = "WHERE 1=1"
            params = []
            
            if date_from:
                base_where += " AND received_time >= ?"
                params.append(date_from)
            
            if date_to:
                base_where += " AND received_time <= ?"
                params.append(date_to)
            
            # Total emails
            cursor.execute(f"SELECT COUNT(*) FROM emails {base_where}", params)
            total_emails = cursor.fetchone()[0]
            
            # Emails by folder
            cursor.execute(f"SELECT folder, COUNT(*) FROM emails {base_where} GROUP BY folder", params)
            emails_by_folder = dict(cursor.fetchall())
            
            # Emails by sender
            cursor.execute(f"SELECT sender_email, COUNT(*) FROM emails {base_where} GROUP BY sender_email ORDER BY COUNT(*) DESC LIMIT 10", params)
            emails_by_sender = dict(cursor.fetchall())
            
            # Daily volume (last 30 days if no date range specified)
            if not date_from and not date_to:
                cursor.execute("""
                SELECT DATE(received_time) as date, COUNT(*) 
                FROM emails 
                WHERE received_time >= date('now', '-30 days')
                GROUP BY DATE(received_time)
                ORDER BY date
                """)
            else:
                cursor.execute(f"""
                SELECT DATE(received_time) as date, COUNT(*) 
                FROM emails {base_where}
                GROUP BY DATE(received_time)
                ORDER BY date
                """, params)
            
            daily_volume = dict(cursor.fetchall())
            
            return {
                'total_emails': total_emails,
                'emails_by_folder': emails_by_folder,
                'emails_by_sender': emails_by_sender,
                'daily_volume': daily_volume,
                'average_response_time': None,  # Would need conversation threading to calculate
                'top_keywords': [],  # Would need text analysis to extract
                'sentiment_distribution': {}  # Would be populated by sentiment analysis
            }
            
        except Exception as e:
            logger.error(f"Error getting email statistics: {str(e)}", exc_info=True)
            return {
                'total_emails': 0,
                'emails_by_folder': {},
                'emails_by_sender': {},
                'daily_volume': {},
                'average_response_time': None,
                'top_keywords': [],
                'sentiment_distribution': {}
            }

    def extract_contacts(self, date_from: Optional[str] = None, date_to: Optional[str] = None, min_interactions: int = 1) -> List[Dict[str, Any]]:
        """
        Extract contact information from emails.
        
        Args:
            date_from (Optional[str]): Start date (ISO format)
            date_to (Optional[str]): End date (ISO format)
            min_interactions (int): Minimum number of interactions required
            
        Returns:
            List[Dict]: List of contact information
        """
        try:
            cursor = self.conn.cursor()
            
            # Build query for sender contacts
            base_where = "WHERE sender_email IS NOT NULL AND sender_email != ''"
            params = []
            
            if date_from:
                base_where += " AND received_time >= ?"
                params.append(date_from)
            
            if date_to:
                base_where += " AND received_time <= ?"
                params.append(date_to)
            
            query = f"""
            SELECT 
                sender_email as email_address,
                sender_name as display_name,
                COUNT(*) as interaction_count,
                MAX(received_time) as last_interaction,
                GROUP_CONCAT(DISTINCT SUBSTR(sender_email, INSTR(sender_email, '@'))) as domains
            FROM emails 
            {base_where}
            GROUP BY sender_email, sender_name
            HAVING COUNT(*) >= ?
            ORDER BY interaction_count DESC
            """
            
            params.append(min_interactions)
            cursor.execute(query, params)
            
            contacts = []
            for row in cursor.fetchall():
                contacts.append({
                    'email_address': row[0],
                    'display_name': row[1] or '',
                    'interaction_count': row[2],
                    'last_interaction': row[3],
                    'domains': [d for d in (row[4] or '').split(',') if d.strip()]
                })
            
            return contacts
            
        except Exception as e:
            logger.error(f"Error extracting contacts: {str(e)}", exc_info=True)
            return []

    def get_emails_by_folder(self, folder_name: str) -> List[Dict[str, Any]]:
        """
        Get all emails from a specific folder.
        
        Args:
            folder_name (str): Name of the folder to retrieve emails from
            
        Returns:
            List[Dict]: List of emails from the specified folder
        """
        try:
            cursor = self.conn.cursor()
            
            query = '''
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
            WHERE folder = ?
            ORDER BY received_time DESC
            '''
            
            cursor.execute(query, (folder_name,))
            return [dict(row) for row in cursor.fetchall()]
            
        except Exception as e:
            logger.error(f"Error getting emails by folder: {str(e)}", exc_info=True)
            return []

    def get_contact_interactions(self, email_address: str) -> List[Dict[str, Any]]:
        """
        Get all email interactions with a specific contact.
        
        Args:
            email_address (str): Email address to get interactions for
            
        Returns:
            List[Dict]: List of email interactions with the contact
        """
        try:
            cursor = self.conn.cursor()
            
            query = '''
            SELECT 
                id,
                folder as Folder,
                subject as Subject,
                received_time as ReceivedTime,
                sent_time as SentOn,
                body as Body,
                CASE 
                    WHEN sender_email = ? THEN 'received'
                    WHEN recipients LIKE ? THEN 'sent'
                    ELSE 'unknown'
                END as interaction_type
            FROM emails 
            WHERE sender_email = ? OR recipients LIKE ?
            ORDER BY COALESCE(sent_time, received_time) DESC
            '''
            
            like_pattern = f"%{email_address}%"
            cursor.execute(query, (email_address, like_pattern, email_address, like_pattern))
            return [dict(row) for row in cursor.fetchall()]
            
        except Exception as e:
            logger.error(f"Error getting contact interactions: {str(e)}", exc_info=True)
            return []

    def get_basic_statistics(self) -> Dict[str, Any]:
        """
        Get basic email statistics for the entire database.
        
        Returns:
            Dict: Basic statistics about all emails
        """
        try:
            cursor = self.conn.cursor()
            
            # Total emails
            cursor.execute("SELECT COUNT(*) FROM emails")
            total_emails = cursor.fetchone()[0]
            
            # Unread emails
            cursor.execute("SELECT COUNT(*) FROM emails WHERE unread = 1")
            unread_count = cursor.fetchone()[0]
            
            # Emails with tasks
            cursor.execute("SELECT COUNT(*) FROM emails WHERE is_task = 1")
            task_count = cursor.fetchone()[0]
            
            # Emails by folder
            cursor.execute("SELECT folder, COUNT(*) FROM emails GROUP BY folder")
            folder_stats = dict(cursor.fetchall())
            
            # Date range
            cursor.execute("SELECT MIN(received_time), MAX(received_time) FROM emails")
            date_range = cursor.fetchone()
            
            return {
                'total_emails': total_emails,
                'unread_count': unread_count,
                'task_count': task_count,
                'folder_distribution': folder_stats,
                'date_range': {
                    'earliest': date_range[0] if date_range[0] else None,
                    'latest': date_range[1] if date_range[1] else None
                }
            }
            
        except Exception as e:
            logger.error(f"Error getting basic statistics: {str(e)}", exc_info=True)
            return {
                'total_emails': 0,
                'unread_count': 0,
                'task_count': 0,
                'folder_distribution': {},
                'date_range': {'earliest': None, 'latest': None}
            }
