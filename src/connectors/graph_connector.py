"""
Microsoft Graph API connector for cross-platform email access.

This connector uses the Microsoft Graph API to access Outlook/Microsoft 365 emails.
It works on any platform (Windows, Mac, Linux) and supports both personal and
organizational accounts.

Requires Azure AD app registration with Mail.Read permission.
"""

import sys
import os
import logging
import re
from datetime import datetime
from typing import List, Optional, Dict, Any

import pytz

from .base import OutlookConnectorBase
from .mailbox_info import MailboxInfo

# Import EmailMetadata from parent
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from EmailMetadata import EmailMetadata

logger = logging.getLogger(__name__)


class GraphAPIConnector(OutlookConnectorBase):
    """
    Microsoft Graph API connector for cross-platform email access.

    Uses the Microsoft Graph API to access emails from Microsoft 365/Outlook.com
    accounts. This requires Azure AD app registration and appropriate permissions.

    Required permissions:
    - Mail.Read (for reading emails)
    - User.Read (for user profile)
    """

    # Graph API base URL
    GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

    def __init__(
        self,
        process_deleted_items: bool = False,
        timezone: Optional[str] = None,
        client_id: Optional[str] = None,
        client_secret: Optional[str] = None,
        tenant_id: Optional[str] = None,
        user_email: Optional[str] = None,
        **kwargs
    ):
        """
        Initialize the Graph API connector.

        Args:
            process_deleted_items: Whether to include Deleted Items folder
            timezone: Local timezone name
            client_id: Azure AD application (client) ID
            client_secret: Azure AD client secret (for app-only auth)
            tenant_id: Azure AD tenant ID
            user_email: User email for delegated auth or mailbox to access
            **kwargs: Additional options
        """
        super().__init__(
            process_deleted_items=process_deleted_items,
            timezone=timezone,
            **kwargs
        )

        # Get credentials from params or environment
        self.client_id = client_id or os.environ.get('GRAPH_CLIENT_ID')
        self.client_secret = client_secret or os.environ.get('GRAPH_CLIENT_SECRET')
        self.tenant_id = tenant_id or os.environ.get('GRAPH_TENANT_ID', 'common')
        self.user_email = user_email or os.environ.get('GRAPH_USER_EMAIL')

        self._access_token = None
        self._token_expiry = None
        self._msal_app = None
        self._authenticated_user = None
        self._use_app_only_auth = bool(client_secret)  # App-only if we have a secret

    @property
    def provider_name(self) -> str:
        return "graph"

    def _get_user_endpoint(self) -> str:
        """
        Get the correct user endpoint path based on auth type.

        For delegated auth (user signs in): /me
        For app-only auth (client credentials): /users/{email}

        Returns:
            API path prefix for user-specific operations
        """
        if self._use_app_only_auth and self.user_email:
            return f"/users/{self.user_email}"
        return "/me"

    @property
    def is_available(self) -> bool:
        """Check if Graph API dependencies are available."""
        try:
            import msal
            import requests
            return self.client_id is not None
        except ImportError:
            return False

    def _get_msal_app(self):
        """Get or create MSAL application instance."""
        if self._msal_app is None:
            try:
                import msal

                if self.client_secret:
                    # Confidential client (app-only or daemon)
                    self._msal_app = msal.ConfidentialClientApplication(
                        self.client_id,
                        authority=f"https://login.microsoftonline.com/{self.tenant_id}",
                        client_credential=self.client_secret,
                    )
                else:
                    # Public client (interactive or device code)
                    self._msal_app = msal.PublicClientApplication(
                        self.client_id,
                        authority=f"https://login.microsoftonline.com/{self.tenant_id}",
                    )
            except Exception as e:
                logger.error(f"Failed to create MSAL app: {e}")
                raise

        return self._msal_app

    def _acquire_token(self) -> str:
        """
        Acquire an access token for Microsoft Graph.

        Returns:
            Access token string

        Raises:
            RuntimeError: If token acquisition fails
        """
        import time

        # Check if we have a valid cached token
        if self._access_token and self._token_expiry:
            if time.time() < self._token_expiry - 300:  # 5 min buffer
                return self._access_token

        try:
            app = self._get_msal_app()
            scopes = ["https://graph.microsoft.com/.default"]

            if self.client_secret:
                # Client credentials flow (app-only)
                result = app.acquire_token_for_client(scopes=scopes)
            else:
                # Try cached token first
                accounts = app.get_accounts()
                if accounts:
                    result = app.acquire_token_silent(scopes=scopes, account=accounts[0])
                else:
                    result = None

                if not result:
                    # Device code flow for interactive auth
                    flow = app.initiate_device_flow(scopes=["Mail.Read", "User.Read"])
                    if "user_code" in flow:
                        logger.info(f"To sign in, visit: {flow['verification_uri']}")
                        logger.info(f"Enter code: {flow['user_code']}")
                        print(f"\n*** AUTHENTICATION REQUIRED ***")
                        print(f"Visit: {flow['verification_uri']}")
                        print(f"Enter code: {flow['user_code']}\n")
                        result = app.acquire_token_by_device_flow(flow)
                    else:
                        raise RuntimeError(f"Device flow initiation failed: {flow.get('error')}")

            if "access_token" in result:
                self._access_token = result["access_token"]
                # Token expires in 'expires_in' seconds
                self._token_expiry = time.time() + result.get("expires_in", 3600)
                return self._access_token
            else:
                error = result.get("error_description", result.get("error", "Unknown error"))
                raise RuntimeError(f"Token acquisition failed: {error}")

        except Exception as e:
            logger.error(f"Failed to acquire token: {e}")
            raise RuntimeError(f"Authentication failed: {e}")

    def _make_request(
        self,
        endpoint: str,
        method: str = "GET",
        params: Optional[Dict] = None,
        body: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """
        Make an authenticated request to Microsoft Graph API.

        Args:
            endpoint: API endpoint (relative to base URL)
            method: HTTP method
            params: Query parameters
            body: Request body (for POST/PATCH)

        Returns:
            Response JSON as dict
        """
        import requests

        token = self._acquire_token()
        url = f"{self.GRAPH_API_BASE}{endpoint}"

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

        try:
            if method == "GET":
                response = requests.get(url, headers=headers, params=params, timeout=30)
            elif method == "POST":
                response = requests.post(url, headers=headers, params=params, json=body, timeout=30)
            else:
                raise ValueError(f"Unsupported method: {method}")

            response.raise_for_status()
            return response.json()

        except requests.exceptions.HTTPError as e:
            logger.error(f"Graph API error: {e.response.status_code} - {e.response.text}")
            raise
        except Exception as e:
            logger.error(f"Request failed: {e}")
            raise

    def _get_current_user(self) -> Dict[str, Any]:
        """Get the current authenticated user's profile."""
        if self._authenticated_user is None:
            try:
                user_endpoint = self._get_user_endpoint()
                self._authenticated_user = self._make_request(user_endpoint)
            except Exception:
                # For app-only auth, use the configured user email
                if self.user_email:
                    self._authenticated_user = {
                        "displayName": self.user_email.split("@")[0],
                        "mail": self.user_email,
                        "userPrincipalName": self.user_email,
                    }
                else:
                    raise

        return self._authenticated_user

    def get_mailboxes(self) -> List[MailboxInfo]:
        """
        Get available mailboxes.

        For Graph API, this returns the authenticated user's mailbox.
        For app-only auth with admin consent, could potentially access multiple mailboxes.
        """
        if not self.is_available:
            return []

        mailboxes = []

        try:
            user = self._get_current_user()
            mailbox = MailboxInfo(
                display_name=user.get("displayName", "Microsoft 365"),
                email_address=user.get("mail") or user.get("userPrincipalName", ""),
                provider=self.provider_name,
                account_type="microsoft365",
                raw_handle={"user_id": user.get("id"), "email": user.get("mail")},
            )
            mailboxes.append(mailbox)

        except Exception as e:
            logger.error(f"Error getting mailboxes: {e}")

        return mailboxes

    def get_mailbox(self, name: str) -> Optional[MailboxInfo]:
        """Get a specific mailbox by name or email."""
        mailboxes = self.get_mailboxes()
        name_lower = name.lower()

        for mailbox in mailboxes:
            if (mailbox.display_name.lower() == name_lower or
                mailbox.email_address.lower() == name_lower):
                return mailbox

        return None

    def _get_folder_id(self, folder_name: str) -> str:
        """
        Get the Graph API folder ID for a folder name.

        Args:
            folder_name: Human-readable folder name

        Returns:
            Graph API well-known folder name or ID
        """
        folder_map = {
            "Inbox": "inbox",
            "Sent Items": "sentitems",
            "Deleted Items": "deleteditems",
            "Drafts": "drafts",
            "Archive": "archive",
        }
        return folder_map.get(folder_name, folder_name.lower())

    @staticmethod
    def clean_email_body(body: str) -> str:
        """Clean email body by removing problematic content."""
        if not body:
            return ""

        body = str(body)

        # Remove HTML tags if present
        body = re.sub(r'<[^>]+>', ' ', body)

        # Remove problematic characters
        body = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', body)
        body = re.sub(r'\r\n|\r|\n', ' ', body)
        body = re.sub(r'\s+', ' ', body)

        # Remove email markers
        body = re.sub(r'From:.*?Sent:.*?(?=\w)', '', body, flags=re.IGNORECASE | re.DOTALL)

        # Escape special characters
        body = body.replace('\\', '\\\\')
        body = body.replace('"', '\\"')
        body = body.replace('\t', ' ')

        return body.strip()

    def _parse_graph_date(self, date_str: str) -> Optional[datetime]:
        """Parse Graph API date string to datetime."""
        if not date_str:
            return None

        try:
            # Graph API returns ISO 8601 format
            # Remove the 'Z' suffix and parse
            date_str = date_str.rstrip("Z")
            if "." in date_str:
                dt = datetime.fromisoformat(date_str)
            else:
                dt = datetime.fromisoformat(date_str)

            # Graph API dates are in UTC
            return dt.replace(tzinfo=pytz.UTC)

        except Exception as e:
            logger.warning(f"Could not parse date: {date_str} - {e}")
            return None

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

        # Parse dates
        start_dt = datetime.fromisoformat(start_date).replace(tzinfo=pytz.UTC)
        end_dt = datetime.fromisoformat(end_date).replace(
            hour=23, minute=59, second=59, tzinfo=pytz.UTC
        )

        # Build date filter for Graph API
        start_iso = start_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        end_iso = end_dt.strftime("%Y-%m-%dT%H:%M:%SZ")

        for mailbox in mailboxes:
            for folder_name in folder_names:
                # Skip Deleted Items if not enabled
                if folder_name == "Deleted Items" and not self.process_deleted_items:
                    continue

                folder_id = self._get_folder_id(folder_name)

                try:
                    # Build the request
                    # Filter by receivedDateTime
                    filter_query = (
                        f"receivedDateTime ge {start_iso} and "
                        f"receivedDateTime le {end_iso}"
                    )

                    # Select specific fields to reduce payload
                    select_fields = (
                        "id,subject,from,toRecipients,receivedDateTime,"
                        "sentDateTime,bodyPreview,body,isRead,categories,"
                        "hasAttachments,importance"
                    )

                    # Paginate through results
                    user_endpoint = self._get_user_endpoint()
                    next_link = f"{user_endpoint}/mailFolders/{folder_id}/messages"
                    params = {
                        "$filter": filter_query,
                        "$select": select_fields,
                        "$orderby": "receivedDateTime desc",
                        "$top": 100,  # Max per page
                    }

                    page_count = 0
                    max_pages = 10  # Limit to 1000 emails per folder

                    while next_link and page_count < max_pages:
                        try:
                            if next_link.startswith("http"):
                                # Full URL from @odata.nextLink
                                import requests
                                token = self._acquire_token()
                                response = requests.get(
                                    next_link,
                                    headers={"Authorization": f"Bearer {token}"},
                                    timeout=30
                                )
                                response.raise_for_status()
                                result = response.json()
                            else:
                                result = self._make_request(next_link, params=params)

                            messages = result.get("value", [])

                            for msg in messages:
                                try:
                                    # Parse dates
                                    received_time = self._parse_graph_date(
                                        msg.get("receivedDateTime")
                                    )
                                    sent_time = self._parse_graph_date(
                                        msg.get("sentDateTime")
                                    )

                                    # Get sender info
                                    from_field = msg.get("from", {}).get("emailAddress", {})
                                    sender_name = from_field.get("name", "")
                                    sender_email = from_field.get("address", "")

                                    # Get recipients
                                    to_recipients = msg.get("toRecipients", [])
                                    to_list = [
                                        r.get("emailAddress", {}).get("address", "")
                                        for r in to_recipients
                                    ]
                                    to_str = "; ".join(filter(None, to_list))

                                    # Get body
                                    body_obj = msg.get("body", {})
                                    body_content = body_obj.get("content", "")
                                    if body_obj.get("contentType") == "html":
                                        body_content = self.clean_email_body(body_content)
                                    else:
                                        body_content = self.clean_email_body(body_content)

                                    # Get categories
                                    categories = msg.get("categories", [])
                                    categories_str = "; ".join(categories)

                                    email_metadata = EmailMetadata(
                                        AccountName=mailbox.display_name,
                                        Entry_ID=msg.get("id", ""),
                                        Folder=folder_name,
                                        Subject=msg.get("subject", ""),
                                        SenderName=sender_name,
                                        SenderEmailAddress=sender_email,
                                        ReceivedTime=received_time or datetime.now(pytz.UTC),
                                        SentOn=sent_time,
                                        To=to_str,
                                        Body=body_content,
                                        Attachments=[],  # Would need separate request
                                        IsMarkedAsTask=False,  # Not directly available
                                        UnRead=not msg.get("isRead", False),
                                        Categories=categories_str,
                                    )
                                    email_data.append(email_metadata)

                                except Exception as e:
                                    logger.debug(f"Error parsing message: {e}")
                                    continue

                            # Check for next page
                            next_link = result.get("@odata.nextLink")
                            params = None  # URL includes params
                            page_count += 1

                        except Exception as e:
                            logger.error(f"Error fetching page: {e}")
                            break

                except Exception as e:
                    logger.error(f"Error getting emails from {folder_name}: {e}")
                    continue

        return email_data
