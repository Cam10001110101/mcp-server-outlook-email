import win32com.client
from datetime import datetime
import pytz
import pywintypes
import re
from EmailMetadata import EmailMetadata
class OutlookConnector:
    def __init__(self, process_deleted_items=False):
        """
        Initialize the Outlook connector.
        
        Args:
            process_deleted_items (bool): Whether to process emails from the Deleted Items folder
        """
        self.process_deleted_items = process_deleted_items
        try:
            self.app = win32com.client.Dispatch("Outlook.Application")
            self.outlook = self.app.GetNamespace("MAPI")
            self.current_user = self.app.Session.CurrentUser
        except Exception as e:
            self.outlook = None
            self.app = None
            self.current_user = None

    def get_mailboxes(self):
        mailboxes = []
        if self.outlook:
            try:
                for account in self.app.Session.Accounts:
                    mailboxes.append(account)
            except Exception:
                pass
        return mailboxes

    def get_mailbox(self, mailbox_name):
        if self.outlook:
            try:
                for account in self.app.Session.Accounts:
                    if account.DisplayName.lower() == mailbox_name.lower():
                        return account
            except Exception:
                pass
        return None

    def _get_store_for_account(self, account):
        """Find the correct MAPI store for a given account by name matching."""
        if not self.outlook:
            return None

        account_name = account.DisplayName.lower()
        account_smtp = getattr(account, 'SmtpAddress', '').lower()

        try:
            for store in self.outlook.Folders:
                store_name = store.Name.lower()

                # Match by display name or email address
                if store_name == account_name:
                    return store
                if account_smtp and (store_name == account_smtp or account_smtp in store_name):
                    return store
                if account_name in store_name or store_name in account_name:
                    return store
        except Exception:
            pass

        # Fallback to DeliveryStore (may work for Exchange)
        try:
            return account.DeliveryStore
        except Exception:
            return None

    def _get_folder_by_name(self, store, folder_name):
        """Get a folder by name when GetDefaultFolder is unavailable."""
        if not store:
            return None

        try:
            for subfolder in store.Folders:
                if subfolder.Name == folder_name:
                    return subfolder
                # Handle common variations
                name_lower = subfolder.Name.lower()
                if folder_name == "Inbox" and name_lower in ['inbox', 'boite de reception']:
                    return subfolder
                if folder_name == "Sent Items" and name_lower in ['sent items', 'sent']:
                    return subfolder
                if folder_name == "Deleted Items" and name_lower in ['deleted items', 'trash']:
                    return subfolder
        except Exception:
            pass
        return None

    def _is_folder_in_store(self, folder, store):
        """Verify a folder belongs to the expected store by comparing StoreIDs."""
        if not folder or not store:
            return False
        try:
            folder_store_id = folder.StoreID
            store_id = store.StoreID
            return folder_store_id == store_id
        except Exception:
            # If we can't verify, assume it's not in the store
            return False

    @staticmethod
    def clean_email_body(body):
        """Clean email body by removing problematic content and escaping for JSON."""
        if not body:
            return ""
            
        # Convert to string if not already
        body = str(body)
        
        # Remove problematic characters and normalize
        body = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', body)  # Remove control characters
        body = re.sub(r'\r\n|\r|\n', ' ', body)  # Normalize line endings to spaces
        body = re.sub(r'\s+', ' ', body)  # Collapse whitespace
        
        # Remove email markers that could break JSON
        body = re.sub(r'From:.*?Sent:.*?(?=\w)', '', body, flags=re.IGNORECASE | re.DOTALL)
        body = re.sub(r'>{2,}.*?(?=\w)', '', body, flags=re.MULTILINE)
        body = re.sub(r'(-{3,}|_{3,}) ?Forwarded message ?(-{3,}|_{3,})', '', body)
        
        # Escape any remaining special characters
        body = body.replace('\\', '\\\\')
        body = body.replace('"', '\\"')
        body = body.replace('\t', ' ')
        
        return body.strip()

    def get_emails_within_date_range(self, folder_names, start_date, end_date, mailboxes):
        email_data = []
        processed_count = 0
        error_count = 0
        skipped_folders = []
        
        # Define folder IDs
        folder_ids = {
            "Inbox": 6,        # olFolderInbox
            "Sent Items": 5,   # olFolderSentMail
            "Deleted Items": 3 # olFolderDeletedItems
        }

        # Convert dates for filtering
        start_datetime = datetime.fromisoformat(start_date)
        end_datetime = datetime.fromisoformat(end_date)
        
        # Convert dates to UTC for comparison
        local_tz = pytz.timezone('America/Chicago')
        start_utc = local_tz.localize(datetime.fromisoformat(start_date).replace(hour=0, minute=0, second=0)).astimezone(pytz.UTC)
        end_utc = local_tz.localize(datetime.fromisoformat(end_date).replace(hour=23, minute=59, second=59)).astimezone(pytz.UTC)
        
        for account in mailboxes:
            try:
                store = self._get_store_for_account(account)
                if store is None:
                    continue

                # Process each requested folder
                for folder_name in folder_names:
                    # Skip Deleted Items folder if not enabled
                    if folder_name == "Deleted Items" and not self.process_deleted_items:
                        continue

                    folder_id = folder_ids.get(folder_name)
                    if not folder_id:
                        continue

                    # Try GetDefaultFolder, verify it's in the correct store, fallback to name lookup
                    folder = None
                    try:
                        temp_folder = store.GetDefaultFolder(folder_id)
                        # Verify the folder belongs to this store (not a different account)
                        if self._is_folder_in_store(temp_folder, store):
                            folder = temp_folder
                    except Exception:
                        pass

                    # Fallback to name lookup if GetDefaultFolder failed or returned wrong store's folder
                    if folder is None:
                        folder = self._get_folder_by_name(store, folder_name)

                    if folder is None:
                        continue
                    if folder and folder.Items.Count > 0:
                        items = folder.Items
                    
                        for i in range(1, items.Count + 1):
                            try:
                                email = items.Item(i)
                                if hasattr(email, 'ReceivedTime'):
                                    # Check if email is within date range
                                    email_time = self.to_utc(email.ReceivedTime)
                                    if not (start_utc <= email_time <= end_utc):
                                        continue
                                    
                                    # Convert dates to UTC datetime objects
                                    received_datetime = self.to_utc(email.ReceivedTime)
                                    sent_datetime = self.to_utc(email.SentOn) if hasattr(email, 'SentOn') and email.SentOn else None
                                    
                                    # Get recipients
                                    to = getattr(email, 'To', '')
                                    if not isinstance(to, str):
                                        to = '; '.join(recipient.Name for recipient in email.Recipients if hasattr(recipient, 'Name'))
                                    
                                    # Get sender email
                                    sender_email = getattr(email, 'SenderEmailAddress', '')
                                    if '/O=EXCHANGELABS/' in sender_email.upper():
                                        # Try to get from Recipients
                                        if hasattr(email, 'Recipients'):
                                            for i in range(1, email.Recipients.Count + 1):
                                                recipient = email.Recipients.Item(i)
                                                if hasattr(recipient, 'Type') and recipient.Type == 1:  # 1 = To
                                                    sender_email = getattr(recipient, 'Address', sender_email)
                                                    break
                                    
                                    # Get attachments
                                    attachments = [attachment.FileName for attachment in email.Attachments] if hasattr(email, 'Attachments') and email.Attachments.Count > 0 else []
                                    
                                    # Clean body
                                    body = self.clean_email_body(getattr(email, 'Body', ''))
                                    
                                    try:
                                        email_metadata = EmailMetadata(
                                            AccountName=account.DisplayName,
                                            Entry_ID=email.EntryID,
                                            Folder=folder_name,
                                            Subject=email.Subject,
                                            SenderName=getattr(email, 'SenderName', ''),
                                            SenderEmailAddress=sender_email,
                                            ReceivedTime=received_datetime,
                                            SentOn=sent_datetime,
                                            To=to,
                                            Body=body,
                                            Attachments=attachments,
                                            IsMarkedAsTask=getattr(email, 'IsMarkedAsTask', False),
                                            UnRead=getattr(email, 'UnRead', False),
                                            Categories='; '.join(email.Categories) if hasattr(email, 'Categories') else ''
                                        )
                                    except Exception:
                                        raise
                                    email_data.append(email_metadata)
                                    processed_count += 1
                            except Exception:
                                error_count += 1
                                continue
                
            except Exception:
                continue

        return email_data

    def to_utc(self, dt):
        try:
            if dt.tzinfo is None or dt.tzinfo.utcoffset(dt) is None:
                local_tz = pytz.timezone('America/Chicago')
                local_dt = local_tz.localize(dt)
                return local_dt.astimezone(pytz.utc)
            else:
                return dt.astimezone(pytz.utc)
        except Exception:
            raise
