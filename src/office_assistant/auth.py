"""Microsoft Graph authentication using MSAL device code flow.

Tokens are cached to ~/.office-assistant/token_cache.json so the user
only needs to authenticate once (refresh tokens last ~90 days).
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import msal
from dotenv import dotenv_values

CACHE_DIR = Path.home() / ".office-assistant"
CACHE_FILE = CACHE_DIR / "token_cache.json"
SCOPES = [
    "Calendars.ReadWrite",
    "Calendars.ReadWrite.Shared",
    "User.Read",
]


def _load_env() -> tuple[str, str]:
    """Load CLIENT_ID and TENANT_ID from the .env file.

    Checks the DOTENV_PATH environment variable first (set by setup.sh
    via the MCP ``-e`` flag), then falls back to ``.env`` in the current
    working directory.
    """
    dotenv_path = os.environ.get("DOTENV_PATH", ".env")
    config = dotenv_values(dotenv_path)

    client_id = config.get("CLIENT_ID", "")
    tenant_id = config.get("TENANT_ID", "")

    if not client_id or not tenant_id:
        raise RuntimeError(
            "CLIENT_ID and TENANT_ID must be set in your .env file. Run /calendar-setup for help."
        )
    return client_id, tenant_id


def _build_cache() -> msal.SerializableTokenCache:
    """Load the persistent token cache from disk."""
    cache = msal.SerializableTokenCache()
    if CACHE_FILE.exists():
        cache.deserialize(CACHE_FILE.read_text())
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    """Write the token cache back to disk if anything changed."""
    if cache.has_state_changed:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        CACHE_FILE.write_text(cache.serialize())
        # Restrict to owner-only read/write since this contains auth tokens.
        CACHE_FILE.chmod(0o600)


def get_token() -> str:
    """Get a valid access token, refreshing or re-authenticating as needed.

    On first use (or when the refresh token expires) this starts an
    interactive device-code flow: the user opens a URL in their browser,
    enters a short code, and signs in with their Microsoft account.

    Returns:
        The access token string.

    Raises:
        RuntimeError: If credentials are missing or authentication fails.
    """
    client_id, tenant_id = _load_env()
    cache = _build_cache()

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
    )

    # Try silent acquisition first (cached / refresh token).
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]

    # Fall back to device-code flow.
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(
            f"Could not start device-code flow: {flow.get('error_description', 'Unknown error')}"
        )

    # Print to stderr because stdout is the MCP stdio transport.
    print(flow["message"], file=sys.stderr, flush=True)

    result = app.acquire_token_by_device_flow(flow)
    _save_cache(cache)

    if "access_token" in result:
        return result["access_token"]

    raise RuntimeError(
        f"Authentication failed: {result.get('error_description', 'Unknown error')}"
    )


def clear_cache() -> bool:
    """Remove the token cache file.

    Returns:
        True if a cache file was deleted, False if none existed.
    """
    if CACHE_FILE.exists():
        CACHE_FILE.unlink()
        return True
    return False
