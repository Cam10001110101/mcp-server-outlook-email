"""
Abstract base class for Outlook email connectors.

All platform-specific connectors (Windows, Mac, Graph) must implement this interface.
"""

from abc import ABC, abstractmethod
from datetime import datetime
from typing import List, Optional
import os

from .mailbox_info import MailboxInfo

# Import EmailMetadata from parent package
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from EmailMetadata import EmailMetadata


class OutlookConnectorBase(ABC):
    """
    Abstract base class for Outlook email connectors.

    This defines the interface that all platform-specific connectors must implement.
    The goal is to provide consistent email access regardless of the underlying
    platform (Windows COM, Mac AppleScript, or Microsoft Graph API).
    """

    def __init__(
        self,
        process_deleted_items: bool = False,
        timezone: Optional[str] = None,
        **kwargs
    ):
        """
        Initialize the connector.

        Args:
            process_deleted_items: Whether to include Deleted Items folder
            timezone: Local timezone name (e.g., 'America/Chicago').
                     If None, uses LOCAL_TIMEZONE env var or defaults to UTC.
            **kwargs: Provider-specific configuration options
        """
        self.process_deleted_items = process_deleted_items
        self.timezone = timezone or os.environ.get('LOCAL_TIMEZONE', 'UTC')
        self._tz_cache = None  # Cache for timezone object

    @property
    @abstractmethod
    def provider_name(self) -> str:
        """Return the provider name (e.g., 'windows', 'mac', 'graph')."""
        pass

    @property
    @abstractmethod
    def is_available(self) -> bool:
        """Check if this connector is available on the current platform."""
        pass

    @abstractmethod
    def get_mailboxes(self) -> List[MailboxInfo]:
        """
        Get all available mailboxes/accounts.

        Returns:
            List of MailboxInfo objects representing available mailboxes.
        """
        pass

    @abstractmethod
    def get_mailbox(self, name: str) -> Optional[MailboxInfo]:
        """
        Get a specific mailbox by name or email address.

        Args:
            name: Display name or email address of the mailbox.

        Returns:
            MailboxInfo if found, None otherwise.
        """
        pass

    @abstractmethod
    def get_emails_within_date_range(
        self,
        folder_names: List[str],
        start_date: str,
        end_date: str,
        mailboxes: List[MailboxInfo]
    ) -> List[EmailMetadata]:
        """
        Retrieve emails within a date range from specified folders and mailboxes.

        Args:
            folder_names: List of folder names to search (e.g., ['Inbox', 'Sent Items'])
            start_date: Start date in ISO format (YYYY-MM-DD)
            end_date: End date in ISO format (YYYY-MM-DD)
            mailboxes: List of MailboxInfo objects to search

        Returns:
            List of EmailMetadata objects representing the retrieved emails.
        """
        pass

    def get_standard_folders(self) -> List[str]:
        """
        Get the list of standard folders to process.

        Returns:
            List of folder names. Includes 'Deleted Items' if process_deleted_items is True.
        """
        folders = ["Inbox", "Sent Items"]
        if self.process_deleted_items:
            folders.append("Deleted Items")
        return folders

    def _parse_date(self, date_str: str) -> datetime:
        """
        Parse an ISO format date string.

        Args:
            date_str: Date in ISO format (YYYY-MM-DD)

        Returns:
            datetime object
        """
        return datetime.fromisoformat(date_str)

    def _get_timezone(self):
        """
        Get the configured timezone object (cached).

        Returns:
            pytz timezone object
        """
        if self._tz_cache is not None:
            return self._tz_cache

        import pytz
        try:
            self._tz_cache = pytz.timezone(self.timezone)
        except Exception:
            self._tz_cache = pytz.UTC
        return self._tz_cache
