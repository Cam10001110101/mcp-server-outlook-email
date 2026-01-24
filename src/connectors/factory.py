"""
Factory for creating platform-appropriate Outlook connectors.

Provides auto-detection of the best available connector and manual selection.
"""

import sys
import os
import logging
from typing import Optional, List, Type

from .base import OutlookConnectorBase
from .mailbox_info import MailboxInfo

logger = logging.getLogger(__name__)


def get_available_providers() -> List[str]:
    """
    Get a list of available providers on the current platform.

    Returns:
        List of provider names that are available (e.g., ['windows'], ['mac', 'graph'])
    """
    available = []

    # Check Windows
    if sys.platform == "win32":
        try:
            import win32com.client
            available.append("windows")
        except ImportError:
            pass

    # Check Mac
    if sys.platform == "darwin":
        # Mac connector uses subprocess, always available on macOS
        available.append("mac")

    # Check Graph (cross-platform)
    try:
        import msal
        import requests
        # Also check for credentials
        if os.environ.get('GRAPH_CLIENT_ID'):
            available.append("graph")
    except ImportError:
        pass

    return available


def detect_best_provider() -> str:
    """
    Auto-detect the best available provider for the current platform.

    Priority:
    1. Windows COM (on Windows) - native, full-featured
    2. Mac AppleScript (on macOS) - native, works offline
    3. Graph API (any platform) - requires Azure setup, works everywhere

    Returns:
        Provider name ('windows', 'mac', or 'graph')

    Raises:
        RuntimeError: If no provider is available
    """
    available = get_available_providers()

    if not available:
        raise RuntimeError(
            "No Outlook connector available. Please ensure:\n"
            "- Windows: Install pywin32 (pip install pywin32)\n"
            "- Mac: Microsoft Outlook for Mac must be installed\n"
            "- Any platform: Configure Graph API credentials (GRAPH_CLIENT_ID, etc.)"
        )

    # Prefer native connectors
    if sys.platform == "win32" and "windows" in available:
        return "windows"
    elif sys.platform == "darwin" and "mac" in available:
        return "mac"
    elif "graph" in available:
        return "graph"
    else:
        # Fall back to first available
        return available[0]


def create_connector(
    provider: str = "auto",
    process_deleted_items: bool = False,
    timezone: Optional[str] = None,
    **kwargs
) -> OutlookConnectorBase:
    """
    Create an Outlook connector instance.

    This factory function creates the appropriate connector based on the
    specified provider or auto-detects the best option for the platform.

    Args:
        provider: Provider to use ('auto', 'windows', 'mac', 'graph')
                 'auto' will detect the best available provider
        process_deleted_items: Whether to include Deleted Items folder
        timezone: Local timezone name (defaults to LOCAL_TIMEZONE env var or UTC)
        **kwargs: Provider-specific options:
            - Windows: (none)
            - Mac: (none)
            - Graph:
                - client_id: Azure AD client ID
                - client_secret: Azure AD client secret
                - tenant_id: Azure AD tenant ID
                - user_email: User email for mailbox access

    Returns:
        OutlookConnectorBase instance

    Raises:
        ValueError: If the specified provider is unknown
        RuntimeError: If the provider is unavailable on the current platform

    Example:
        # Auto-detect best provider
        connector = create_connector()

        # Explicitly use Graph API
        connector = create_connector(
            provider="graph",
            client_id="your-client-id",
            client_secret="your-secret",
            tenant_id="your-tenant"
        )

        # Use Windows COM with custom timezone
        connector = create_connector(
            provider="windows",
            timezone="America/New_York",
            process_deleted_items=True
        )
    """
    # Auto-detect if needed
    if provider == "auto":
        provider = detect_best_provider()
        logger.info(f"Auto-detected provider: {provider}")

    # Get common args
    common_args = {
        "process_deleted_items": process_deleted_items,
        "timezone": timezone,
    }

    # Create the appropriate connector
    if provider == "windows":
        from .windows_connector import WindowsOutlookConnector

        if sys.platform != "win32":
            raise RuntimeError("Windows connector is only available on Windows")

        try:
            import win32com.client
        except ImportError:
            raise RuntimeError(
                "pywin32 is not installed. Install with: pip install pywin32"
            )

        return WindowsOutlookConnector(**common_args, **kwargs)

    elif provider == "mac":
        from .mac_connector import MacOutlookConnector

        if sys.platform != "darwin":
            raise RuntimeError("Mac connector is only available on macOS")

        connector = MacOutlookConnector(**common_args, **kwargs)

        if not connector.is_available:
            raise RuntimeError(
                "Microsoft Outlook for Mac is not installed or not accessible. "
                "Please install Outlook or try the Graph API connector."
            )

        return connector

    elif provider == "graph":
        from .graph_connector import GraphAPIConnector

        try:
            import msal
            import requests
        except ImportError:
            raise RuntimeError(
                "Graph API dependencies not installed. Install with: "
                "pip install msal requests"
            )

        # Get Graph-specific credentials
        graph_args = {
            "client_id": kwargs.get("client_id") or os.environ.get("GRAPH_CLIENT_ID"),
            "client_secret": kwargs.get("client_secret") or os.environ.get("GRAPH_CLIENT_SECRET"),
            "tenant_id": kwargs.get("tenant_id") or os.environ.get("GRAPH_TENANT_ID"),
            "user_email": kwargs.get("user_email") or os.environ.get("GRAPH_USER_EMAIL"),
        }

        if not graph_args["client_id"]:
            raise RuntimeError(
                "Graph API client_id is required. Set GRAPH_CLIENT_ID environment "
                "variable or pass client_id parameter."
            )

        return GraphAPIConnector(**common_args, **graph_args)

    else:
        available = get_available_providers()
        raise ValueError(
            f"Unknown provider: '{provider}'. "
            f"Available providers: {['auto'] + available}"
        )


# Convenience function for backward compatibility
def get_connector(**kwargs) -> OutlookConnectorBase:
    """Alias for create_connector() for backward compatibility."""
    return create_connector(**kwargs)
