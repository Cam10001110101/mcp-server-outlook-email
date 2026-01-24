"""
Outlook Email Connectors Package

Provides platform-agnostic email access through multiple providers:
- Windows: COM automation via pywin32
- Mac: AppleScript via osascript subprocess
- Graph: Microsoft Graph API (cross-platform)
"""

from .base import OutlookConnectorBase
from .mailbox_info import MailboxInfo
from .factory import create_connector, get_available_providers

__all__ = [
    "OutlookConnectorBase",
    "MailboxInfo",
    "create_connector",
    "get_available_providers",
]
