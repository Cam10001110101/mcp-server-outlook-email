"""
Platform-agnostic mailbox information dataclass.

This provides a consistent interface for mailbox data across all providers
(Windows COM, Mac AppleScript, Microsoft Graph API).
"""

from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class MailboxInfo:
    """
    Platform-agnostic representation of an email mailbox/account.

    Attributes:
        display_name: Human-readable name of the mailbox (e.g., "Work Account")
        email_address: Email address associated with the mailbox
        provider: Provider type ('windows', 'mac', 'graph')
        account_type: Type of account (e.g., 'exchange', 'imap', 'microsoft365')
        raw_handle: Platform-specific object for internal use (COM Account, etc.)
    """
    display_name: str
    email_address: str
    provider: str
    account_type: str = "unknown"
    raw_handle: Any = field(default=None, repr=False)

    def __str__(self) -> str:
        return f"{self.display_name} <{self.email_address}>"

    def __hash__(self) -> int:
        return hash((self.display_name.lower(), self.email_address.lower()))

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, MailboxInfo):
            return False
        return (
            self.display_name.lower() == other.display_name.lower() and
            self.email_address.lower() == other.email_address.lower()
        )

    @property
    def is_exchange(self) -> bool:
        """Check if this is an Exchange/Microsoft 365 account."""
        return self.account_type.lower() in ('exchange', 'microsoft365', 'office365')

    def to_dict(self) -> dict:
        """Convert to dictionary (excluding raw_handle)."""
        return {
            "display_name": self.display_name,
            "email_address": self.email_address,
            "provider": self.provider,
            "account_type": self.account_type,
        }
