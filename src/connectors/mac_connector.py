"""
Mac Outlook connector using AppleScript via osascript subprocess.

This connector interfaces with Microsoft Outlook for Mac using AppleScript.
Works with the classic Outlook for Mac (which has AppleScript support).
The "New Outlook" for Mac may have limited AppleScript support.
"""

import sys
import subprocess
import json
import logging
import re
from datetime import datetime
from typing import List, Optional, Dict, Any

import pytz

from .base import OutlookConnectorBase
from .mailbox_info import MailboxInfo

# Import EmailMetadata from parent
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from EmailMetadata import EmailMetadata

logger = logging.getLogger(__name__)


class MacOutlookConnector(OutlookConnectorBase):
    """
    Mac Outlook connector using AppleScript automation.

    Uses osascript subprocess to execute AppleScript commands that interact
    with Microsoft Outlook for Mac. This requires Outlook to be installed
    and the user to have granted automation permissions.
    """

    # AppleScript delimiter for parsing output
    DELIMITER = "|||"

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

        self._outlook_available = None
        self._is_new_outlook = None
        self._tz_cache = None  # Cache timezone object

    @property
    def provider_name(self) -> str:
        return "mac"

    def _escape_applescript(self, value: str) -> str:
        """
        Escape a string for safe interpolation into AppleScript.

        Handles quotes and backslashes that could break AppleScript syntax.
        """
        if not value:
            return ""
        # Escape backslashes first, then quotes
        return value.replace('\\', '\\\\').replace('"', '\\"')

    @property
    def is_available(self) -> bool:
        """Check if Mac Outlook with AppleScript is available."""
        if sys.platform != "darwin":
            return False

        if self._outlook_available is not None:
            return self._outlook_available

        # Check if Outlook is installed and AppleScript accessible
        script = '''
        tell application "System Events"
            return (exists application process "Microsoft Outlook") or (exists application file id "com.microsoft.Outlook")
        end tell
        '''
        try:
            result = self._run_applescript(script)
            self._outlook_available = result.strip().lower() == "true"
        except Exception:
            # Try alternative check
            try:
                result = subprocess.run(
                    ["osascript", "-e", 'tell application "Microsoft Outlook" to name'],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                self._outlook_available = result.returncode == 0
            except Exception:
                self._outlook_available = False

        return self._outlook_available

    def _run_applescript(self, script: str, timeout: int = 30) -> str:
        """
        Execute an AppleScript and return the output.

        Args:
            script: AppleScript code to execute
            timeout: Timeout in seconds

        Returns:
            Script output as string

        Raises:
            RuntimeError: If script execution fails
        """
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=timeout
            )

            if result.returncode != 0:
                error_msg = result.stderr.strip() or "Unknown AppleScript error"
                logger.error(f"AppleScript error: {error_msg}")
                raise RuntimeError(f"AppleScript failed: {error_msg}")

            return result.stdout.strip()

        except subprocess.TimeoutExpired:
            raise RuntimeError(f"AppleScript timed out after {timeout} seconds")
        except Exception as e:
            raise RuntimeError(f"AppleScript execution failed: {e}")

    def _check_new_outlook(self) -> bool:
        """Check if the user is running the New Outlook for Mac."""
        if self._is_new_outlook is not None:
            return self._is_new_outlook

        script = '''
        tell application "Microsoft Outlook"
            try
                -- New Outlook uses different internal structure
                return (version starts with "16.9") or (version > "16.80")
            on error
                return false
            end try
        end tell
        '''
        try:
            result = self._run_applescript(script)
            self._is_new_outlook = result.strip().lower() == "true"
        except Exception:
            self._is_new_outlook = False

        if self._is_new_outlook:
            logger.warning(
                "New Outlook for Mac detected. AppleScript support may be limited. "
                "Consider using the Graph API connector instead."
            )

        return self._is_new_outlook

    def get_mailboxes(self) -> List[MailboxInfo]:
        """Get all available mailboxes from Outlook for Mac."""
        if not self.is_available:
            return []

        self._check_new_outlook()

        # First try to get accounts via traditional API
        script = f'''
        tell application "Microsoft Outlook"
            set accountList to {{}}
            repeat with acct in exchange accounts
                set acctInfo to (name of acct) & "{self.DELIMITER}" & (email address of acct) & "{self.DELIMITER}" & "exchange"
                set end of accountList to acctInfo
            end repeat
            repeat with acct in imap accounts
                set acctInfo to (name of acct) & "{self.DELIMITER}" & (email address of acct) & "{self.DELIMITER}" & "imap"
                set end of accountList to acctInfo
            end repeat
            repeat with acct in pop accounts
                set acctInfo to (name of acct) & "{self.DELIMITER}" & (email address of acct) & "{self.DELIMITER}" & "pop"
                set end of accountList to acctInfo
            end repeat
            return accountList
        end tell
        '''

        mailboxes = []
        try:
            result = self._run_applescript(script)
            if result:
                # Parse AppleScript list output
                # Format: "name|||email|||type, name|||email|||type, ..."
                for entry in result.split(", "):
                    parts = entry.strip().split(self.DELIMITER)
                    if len(parts) >= 3:
                        mailbox = MailboxInfo(
                            display_name=parts[0],
                            email_address=parts[1],
                            provider=self.provider_name,
                            account_type=parts[2],
                            raw_handle={"name": parts[0], "type": parts[2]}
                        )
                        mailboxes.append(mailbox)
        except Exception as e:
            logger.error(f"Error getting mailboxes via accounts: {e}")

        # If no accounts found, try accessing mail folders directly
        # This handles New Outlook for Mac and local "On My Computer" folders
        if not mailboxes:
            logger.info("No accounts found via API, checking for direct mail folder access")
            try:
                # Check if we can access mail folders directly
                folder_check = '''
                tell application "Microsoft Outlook"
                    set folderCount to count of mail folders
                    return folderCount
                end tell
                '''
                folder_count = self._run_applescript(folder_check)
                if folder_count and int(folder_count) > 0:
                    # Create a synthetic "local" mailbox for direct folder access
                    mailbox = MailboxInfo(
                        display_name="On My Computer",
                        email_address="local@localhost",
                        provider=self.provider_name,
                        account_type="local",
                        raw_handle={"name": "On My Computer", "type": "local", "direct_folders": True}
                    )
                    mailboxes.append(mailbox)
                    logger.info(f"Found {folder_count} mail folders via direct access")
            except Exception as e:
                logger.error(f"Error checking direct mail folders: {e}")

        return mailboxes

    def get_mailbox(self, name: str) -> Optional[MailboxInfo]:
        """Get a specific mailbox by name or email address."""
        mailboxes = self.get_mailboxes()
        name_lower = name.lower()

        for mailbox in mailboxes:
            if (mailbox.display_name.lower() == name_lower or
                mailbox.email_address.lower() == name_lower):
                return mailbox

        return None

    @staticmethod
    def clean_email_body(body: str) -> str:
        """Clean email body by removing problematic content."""
        if not body:
            return ""

        body = str(body)

        # Remove problematic characters
        body = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', body)
        body = re.sub(r'\r\n|\r|\n', ' ', body)
        body = re.sub(r'\s+', ' ', body)

        # Remove email markers
        body = re.sub(r'From:.*?Sent:.*?(?=\w)', '', body, flags=re.IGNORECASE | re.DOTALL)
        body = re.sub(r'>{2,}.*?(?=\w)', '', body, flags=re.MULTILINE)
        body = re.sub(r'(-{3,}|_{3,}) ?Forwarded message ?(-{3,}|_{3,})', '', body)

        # Escape special characters
        body = body.replace('\\', '\\\\')
        body = body.replace('"', '\\"')
        body = body.replace('\t', ' ')

        return body.strip()

    def _parse_applescript_date(self, date_str: str) -> Optional[datetime]:
        """
        Parse AppleScript date output to datetime.

        AppleScript dates are typically in format: "day, Month DD, YYYY at HH:MM:SS AM/PM"
        """
        if not date_str or date_str == "missing value":
            return None

        # Common formats from AppleScript
        formats = [
            "%A, %B %d, %Y at %I:%M:%S %p",  # Wednesday, January 15, 2025 at 10:30:00 AM
            "%m/%d/%Y %H:%M:%S",              # 01/15/2025 10:30:00
            "%Y-%m-%dT%H:%M:%S",              # ISO format
            "%B %d, %Y at %I:%M:%S %p",       # January 15, 2025 at 10:30:00 AM
        ]

        for fmt in formats:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except ValueError:
                continue

        logger.warning(f"Could not parse date: {date_str}")
        return None

    def _to_utc(self, dt: Optional[datetime]) -> Optional[datetime]:
        """Convert datetime to UTC."""
        if dt is None:
            return None

        try:
            if dt.tzinfo is None or dt.tzinfo.utcoffset(dt) is None:
                local_tz = self._get_timezone()
                local_dt = local_tz.localize(dt)
                return local_dt.astimezone(pytz.utc)
            else:
                return dt.astimezone(pytz.utc)
        except Exception as e:
            logger.error(f"Error converting to UTC: {e}")
            return dt

    def _get_folder_script_name(self, folder_name: str, account_type: str) -> str:
        """Get the AppleScript folder name for a given folder."""
        folder_map = {
            "Inbox": "inbox",
            "Sent Items": "sent items",
            "Deleted Items": "deleted items",
        }
        return folder_map.get(folder_name, folder_name.lower())

    def get_emails_within_date_range(
        self,
        folder_names: List[str],
        start_date: str,
        end_date: str,
        mailboxes: List[MailboxInfo]
    ) -> List[EmailMetadata]:
        """Retrieve emails within a date range from specified folders."""
        if not self.is_available:
            return []

        email_data = []
        local_tz = self._get_timezone()

        # Parse dates
        start_dt = local_tz.localize(
            datetime.fromisoformat(start_date).replace(hour=0, minute=0, second=0)
        ).astimezone(pytz.UTC)
        end_dt = local_tz.localize(
            datetime.fromisoformat(end_date).replace(hour=23, minute=59, second=59)
        ).astimezone(pytz.UTC)

        for mailbox in mailboxes:
            raw_handle = mailbox.raw_handle or {}
            account_name = raw_handle.get("name", mailbox.display_name)
            account_type = raw_handle.get("type", mailbox.account_type)
            direct_folders = raw_handle.get("direct_folders", False)

            for folder_name in folder_names:
                # Skip Deleted Items if not enabled
                if folder_name == "Deleted Items" and not self.process_deleted_items:
                    continue

                folder_script_name = self._get_folder_script_name(folder_name, account_type)

                # Escape names for safe AppleScript interpolation
                safe_folder_name = self._escape_applescript(folder_name)
                safe_account_name = self._escape_applescript(account_name)

                # Choose script based on whether we're using direct folder access
                if direct_folders:
                    # Direct mail folder access (for New Outlook / On My Computer)
                    script = f'''
                    tell application "Microsoft Outlook"
                        set emailList to {{}}

                        -- Find the folder directly from mail folders
                        repeat with f in mail folders
                            if name of f is "{safe_folder_name}" then
                                try
                                    set msgs to messages of f

                                    repeat with msg in msgs
                                        try
                                            set msgDate to time received of msg

                                            set msgData to (id of msg as string) & "{self.DELIMITER}"
                                            set msgData to msgData & (subject of msg as string) & "{self.DELIMITER}"

                                            -- Handle sender safely
                                            try
                                                set msgData to msgData & (name of sender of msg as string) & "{self.DELIMITER}"
                                                set msgData to msgData & (address of sender of msg as string) & "{self.DELIMITER}"
                                            on error
                                                set msgData to msgData & "Unknown{self.DELIMITER}unknown@local{self.DELIMITER}"
                                            end try

                                            set msgData to msgData & (msgDate as string) & "{self.DELIMITER}"

                                            -- Get recipients
                                            set recipList to ""
                                            try
                                                repeat with recip in to recipients of msg
                                                    set recipList to recipList & (address of recip) & ";"
                                                end repeat
                                            end try
                                            set msgData to msgData & recipList & "{self.DELIMITER}"

                                            -- Get plain text content
                                            try
                                                set msgData to msgData & (plain text content of msg as string) & "{self.DELIMITER}"
                                            on error
                                                set msgData to msgData & "{self.DELIMITER}"
                                            end try

                                            -- Check if unread
                                            try
                                                set msgData to msgData & (is read of msg as string) & "{self.DELIMITER}"
                                            on error
                                                set msgData to msgData & "true{self.DELIMITER}"
                                            end try

                                            -- Get categories
                                            set catList to ""
                                            try
                                                repeat with cat in categories of msg
                                                    set catList to catList & (name of cat) & ";"
                                                end repeat
                                            end try
                                            set msgData to msgData & catList

                                            set end of emailList to msgData

                                            if (count of emailList) > 500 then
                                                exit repeat
                                            end if
                                        on error
                                            -- Skip problematic messages
                                        end try
                                    end repeat
                                on error errMsg
                                    -- Error accessing folder
                                end try
                                exit repeat
                            end if
                        end repeat

                        return emailList
                    end tell
                    '''
                else:
                    # Traditional account-based access
                    script = f'''
                    tell application "Microsoft Outlook"
                        set emailList to {{}}

                        -- Find the account
                        set targetAccount to missing value
                        repeat with acct in exchange accounts
                            if name of acct is "{safe_account_name}" then
                                set targetAccount to acct
                                exit repeat
                            end if
                        end repeat

                        if targetAccount is missing value then
                            repeat with acct in imap accounts
                                if name of acct is "{safe_account_name}" then
                                    set targetAccount to acct
                                    exit repeat
                                end if
                            end repeat
                        end if

                        if targetAccount is not missing value then
                            try
                                set targetFolder to {folder_script_name} of targetAccount
                                set msgs to messages of targetFolder

                                repeat with msg in msgs
                                    try
                                        set msgDate to time received of msg

                                        -- Basic date filtering (AppleScript date comparison)
                                        set msgData to (id of msg as string) & "{self.DELIMITER}"
                                        set msgData to msgData & (subject of msg as string) & "{self.DELIMITER}"
                                        set msgData to msgData & (name of sender of msg as string) & "{self.DELIMITER}"
                                        set msgData to msgData & (address of sender of msg as string) & "{self.DELIMITER}"
                                        set msgData to msgData & (msgDate as string) & "{self.DELIMITER}"

                                        -- Get recipients
                                        set recipList to ""
                                        repeat with recip in to recipients of msg
                                            set recipList to recipList & (address of recip) & ";"
                                        end repeat
                                        set msgData to msgData & recipList & "{self.DELIMITER}"

                                        -- Get plain text content
                                        set msgData to msgData & (plain text content of msg as string) & "{self.DELIMITER}"

                                        -- Check if unread
                                        set msgData to msgData & (is read of msg as string) & "{self.DELIMITER}"

                                        -- Get categories
                                        set catList to ""
                                        try
                                            repeat with cat in categories of msg
                                                set catList to catList & (name of cat) & ";"
                                            end repeat
                                        end try
                                        set msgData to msgData & catList

                                        set end of emailList to msgData

                                        -- Limit to prevent memory issues
                                        if (count of emailList) > 500 then
                                            exit repeat
                                        end if
                                    on error
                                        -- Skip problematic messages
                                    end try
                                end repeat
                            on error errMsg
                                -- Folder may not exist
                            end try
                        end if

                        return emailList
                    end tell
                    '''

                try:
                    result = self._run_applescript(script, timeout=120)
                    if result:
                        # Parse the result
                        for entry in result.split(", "):
                            entry = entry.strip()
                            if not entry:
                                continue

                            # Split with maxsplit to handle bodies containing delimiter
                            # Fields: id, subject, sender_name, sender_email, date, recipients, body, is_read, categories
                            # We split into max 9 parts; if body contains "|||", extra parts go to categories
                            parts = entry.split(self.DELIMITER, maxsplit=8)
                            if len(parts) < 7:
                                logger.debug(f"Skipping entry with insufficient fields: {len(parts)}")
                                continue

                            try:
                                # Parse date and check range
                                received_time = self._parse_applescript_date(parts[4])
                                if received_time:
                                    received_utc = self._to_utc(received_time)
                                    if received_utc and not (start_dt <= received_utc <= end_dt):
                                        continue
                                else:
                                    # Log when date parsing fails instead of silently using now()
                                    logger.debug(f"Could not parse date, using current time for email: {parts[1][:50]}")
                                    received_utc = datetime.now(pytz.UTC)

                                # Parse fields safely
                                is_read = parts[7].lower() == "true" if len(parts) > 7 else False
                                categories = parts[8] if len(parts) > 8 else ""

                                email_metadata = EmailMetadata(
                                    AccountName=mailbox.display_name,
                                    Entry_ID=parts[0],
                                    Folder=folder_name,
                                    Subject=parts[1],
                                    SenderName=parts[2],
                                    SenderEmailAddress=parts[3],
                                    ReceivedTime=received_utc,
                                    SentOn=received_utc,  # AppleScript may not distinguish
                                    To=parts[5].rstrip(";"),
                                    Body=self.clean_email_body(parts[6]),
                                    Attachments=[],  # Would need additional AppleScript
                                    IsMarkedAsTask=False,
                                    UnRead=not is_read,
                                    Categories=categories.rstrip(";")
                                )
                                email_data.append(email_metadata)

                            except Exception as e:
                                logger.debug(f"Error parsing email entry: {e}")
                                continue

                except Exception as e:
                    logger.error(f"Error retrieving emails from {folder_name}: {e}")
                    continue

        return email_data
