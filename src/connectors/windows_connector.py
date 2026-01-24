"""
Windows Outlook connector using COM automation via pywin32.

This is a refactored version of the original OutlookConnector.py that implements
the OutlookConnectorBase interface for cross-platform compatibility.
"""

import sys
import re
import logging
from datetime import datetime
from typing import List, Optional

import pytz

from .base import OutlookConnectorBase
from .mailbox_info import MailboxInfo

# Import EmailMetadata from parent
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from EmailMetadata import EmailMetadata

logger = logging.getLogger(__name__)


class WindowsOutlookConnector(OutlookConnectorBase):
    """
    Windows Outlook connector using COM automation.

    Uses win32com.client to interface with Microsoft Outlook via COM.
    Only available on Windows platforms with Outlook installed.
    """

    def __init__(
        self,
        process_deleted_items: bool = False,
        timezone: Optional[str] = None,
        **kwargs
    ):
        super().__init__(
            process_deleted_items=process_deleted_items,
            timezone=timezone,
            **kwargs
        )

        self.app = None
        self.outlook = None
        self.current_user = None

        if self.is_available:
            self._initialize_com()

    def _initialize_com(self) -> None:
        """Initialize COM connection to Outlook."""
        try:
            import win32com.client
            self.app = win32com.client.Dispatch("Outlook.Application")
            self.outlook = self.app.GetNamespace("MAPI")
            self.current_user = self.app.Session.CurrentUser
            logger.info("Windows Outlook COM connection established")
        except Exception as e:
            logger.error(f"Failed to initialize Outlook COM: {e}")
            self.app = None
            self.outlook = None
            self.current_user = None

    @property
    def provider_name(self) -> str:
        return "windows"

    @property
    def is_available(self) -> bool:
        """Check if Windows COM automation is available."""
        if sys.platform != "win32":
            return False
        try:
            import win32com.client
            return True
        except ImportError:
            return False

    def get_mailboxes(self) -> List[MailboxInfo]:
        """Get all available mailboxes from Outlook."""
        mailboxes = []
        if not self.outlook:
            return mailboxes

        try:
            for account in self.app.Session.Accounts:
                mailbox = MailboxInfo(
                    display_name=account.DisplayName,
                    email_address=getattr(account, 'SmtpAddress', account.DisplayName),
                    provider=self.provider_name,
                    account_type=self._get_account_type(account),
                    raw_handle=account
                )
                mailboxes.append(mailbox)
        except Exception as e:
            logger.error(f"Error getting mailboxes: {e}")

        return mailboxes

    def get_mailbox(self, name: str) -> Optional[MailboxInfo]:
        """Get a specific mailbox by name or email address."""
        if not self.outlook:
            return None

        name_lower = name.lower()
        try:
            for account in self.app.Session.Accounts:
                if (account.DisplayName.lower() == name_lower or
                    getattr(account, 'SmtpAddress', '').lower() == name_lower):
                    return MailboxInfo(
                        display_name=account.DisplayName,
                        email_address=getattr(account, 'SmtpAddress', account.DisplayName),
                        provider=self.provider_name,
                        account_type=self._get_account_type(account),
                        raw_handle=account
                    )
        except Exception as e:
            logger.error(f"Error getting mailbox '{name}': {e}")

        return None

    def _get_account_type(self, account) -> str:
        """Determine the account type from the COM Account object."""
        try:
            account_type = getattr(account, 'AccountType', None)
            if account_type == 1:  # olExchange
                return "exchange"
            elif account_type == 3:  # olPop3
                return "pop3"
            elif account_type == 4:  # olImap
                return "imap"
            elif account_type == 5:  # olHttp (Office 365/Outlook.com)
                return "microsoft365"
        except Exception:
            pass
        return "unknown"

    def _get_store_for_account(self, account) -> Optional[object]:
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

    def _get_folder_by_name(self, store, folder_name: str) -> Optional[object]:
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

    def _is_folder_in_store(self, folder, store) -> bool:
        """Verify a folder belongs to the expected store by comparing StoreIDs."""
        if not folder or not store:
            return False
        try:
            return folder.StoreID == store.StoreID
        except Exception:
            return False

    @staticmethod
    def clean_email_body(body: str) -> str:
        """Clean email body by removing problematic content."""
        if not body:
            return ""

        body = str(body)

        # Remove problematic characters and normalize
        body = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', body)
        body = re.sub(r'\r\n|\r|\n', ' ', body)
        body = re.sub(r'\s+', ' ', body)

        # Remove email markers that could break JSON
        body = re.sub(r'From:.*?Sent:.*?(?=\w)', '', body, flags=re.IGNORECASE | re.DOTALL)
        body = re.sub(r'>{2,}.*?(?=\w)', '', body, flags=re.MULTILINE)
        body = re.sub(r'(-{3,}|_{3,}) ?Forwarded message ?(-{3,}|_{3,})', '', body)

        # Escape special characters
        body = body.replace('\\', '\\\\')
        body = body.replace('"', '\\"')
        body = body.replace('\t', ' ')

        return body.strip()

    def to_utc(self, dt) -> datetime:
        """Convert a datetime to UTC."""
        try:
            if dt.tzinfo is None or dt.tzinfo.utcoffset(dt) is None:
                local_tz = self._get_timezone()
                local_dt = local_tz.localize(dt)
                return local_dt.astimezone(pytz.utc)
            else:
                return dt.astimezone(pytz.utc)
        except Exception:
            raise

    def get_emails_within_date_range(
        self,
        folder_names: List[str],
        start_date: str,
        end_date: str,
        mailboxes: List[MailboxInfo]
    ) -> List[EmailMetadata]:
        """Retrieve emails within a date range from specified folders."""
        email_data = []

        # Define folder IDs
        folder_ids = {
            "Inbox": 6,        # olFolderInbox
            "Sent Items": 5,   # olFolderSentMail
            "Deleted Items": 3 # olFolderDeletedItems
        }

        # Convert dates for filtering
        local_tz = self._get_timezone()
        start_utc = local_tz.localize(
            datetime.fromisoformat(start_date).replace(hour=0, minute=0, second=0)
        ).astimezone(pytz.UTC)
        end_utc = local_tz.localize(
            datetime.fromisoformat(end_date).replace(hour=23, minute=59, second=59)
        ).astimezone(pytz.UTC)

        for mailbox_info in mailboxes:
            # Get the raw COM account object
            account = mailbox_info.raw_handle
            if account is None:
                continue

            try:
                store = self._get_store_for_account(account)
                if store is None:
                    continue

                for folder_name in folder_names:
                    # Skip Deleted Items if not enabled
                    if folder_name == "Deleted Items" and not self.process_deleted_items:
                        continue

                    folder_id = folder_ids.get(folder_name)
                    if not folder_id:
                        continue

                    # Try GetDefaultFolder, verify it's in the correct store
                    folder = None
                    try:
                        temp_folder = store.GetDefaultFolder(folder_id)
                        if self._is_folder_in_store(temp_folder, store):
                            folder = temp_folder
                    except Exception:
                        pass

                    # Fallback to name lookup
                    if folder is None:
                        folder = self._get_folder_by_name(store, folder_name)

                    if folder is None:
                        continue

                    if folder.Items.Count > 0:
                        items = folder.Items

                        for i in range(1, items.Count + 1):
                            try:
                                email = items.Item(i)
                                if hasattr(email, 'ReceivedTime'):
                                    # Check date range
                                    email_time = self.to_utc(email.ReceivedTime)
                                    if not (start_utc <= email_time <= end_utc):
                                        continue

                                    # Convert dates
                                    received_datetime = self.to_utc(email.ReceivedTime)
                                    sent_datetime = self.to_utc(email.SentOn) if hasattr(email, 'SentOn') and email.SentOn else None

                                    # Get recipients
                                    to = getattr(email, 'To', '')
                                    if not isinstance(to, str):
                                        to = '; '.join(r.Name for r in email.Recipients if hasattr(r, 'Name'))

                                    # Get sender email
                                    sender_email = getattr(email, 'SenderEmailAddress', '')
                                    if '/O=EXCHANGELABS/' in sender_email.upper():
                                        if hasattr(email, 'Recipients'):
                                            for j in range(1, email.Recipients.Count + 1):
                                                recipient = email.Recipients.Item(j)
                                                if hasattr(recipient, 'Type') and recipient.Type == 1:
                                                    sender_email = getattr(recipient, 'Address', sender_email)
                                                    break

                                    # Get attachments
                                    attachments = []
                                    if hasattr(email, 'Attachments') and email.Attachments.Count > 0:
                                        attachments = [att.FileName for att in email.Attachments]

                                    # Clean body
                                    body = self.clean_email_body(getattr(email, 'Body', ''))

                                    email_metadata = EmailMetadata(
                                        AccountName=mailbox_info.display_name,
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
                                    email_data.append(email_metadata)

                            except Exception as e:
                                logger.debug(f"Error processing email: {e}")
                                continue

            except Exception as e:
                logger.error(f"Error processing mailbox {mailbox_info.display_name}: {e}")
                continue

        return email_data
