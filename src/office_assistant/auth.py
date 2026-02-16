"""Microsoft Graph authentication using MSAL device code flow.

Tokens are cached to ~/.office-assistant/token_cache.json so the user
only needs to authenticate once (refresh tokens last ~90 days).

Supports both organisational (work/school) and personal Microsoft accounts.
Set ``TENANT_ID=consumers`` in ``.env`` for personal accounts, or use your
Azure AD tenant ID for organisational accounts.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import msal
from dotenv import dotenv_values

CACHE_DIR = Path.home() / ".office-assistant"
CACHE_FILE = CACHE_DIR / "token_cache.json"

# Organisational accounts support shared calendar access; personal accounts
# do not, so we request a narrower set of scopes for consumer tenants.
_ORG_SCOPES = [
    "Calendars.ReadWrite",
    "Calendars.ReadWrite.Shared",
    "Place.Read.All",
    "User.Read",
]
_PERSONAL_SCOPES = [
    "Calendars.ReadWrite",
    "User.Read",
]

# Tenant IDs that indicate a personal / multi-tenant authority rather than
# a specific Azure AD organisation.
_NON_ORG_TENANTS = {"consumers", "common"}


def _is_personal_tenant(tenant_id: str) -> bool:
    """Return True if *tenant_id* targets personal Microsoft accounts."""
    return tenant_id.lower() in _NON_ORG_TENANTS


def _load_env() -> tuple[str, str]:
    """Load CLIENT_ID and TENANT_ID from the .env file.

    Checks the DOTENV_PATH environment variable first (set by setup.sh
    via the MCP ``-e`` flag), then falls back to ``.env`` in the current
    working directory.

    For personal Microsoft accounts set ``TENANT_ID=consumers``.
    For organisational (work/school) accounts use your Azure AD tenant ID.
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
        try:
            cache.deserialize(CACHE_FILE.read_text())
        except Exception:
            # Treat unreadable/corrupt cache as empty so auth can recover.
            return msal.SerializableTokenCache()
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

    # tenant_id can be:
    #   - A specific tenant GUID: restricts sign-in to that org's accounts
    #   - "common": allows both work/school and personal Microsoft accounts
    #   - "consumers": allows personal Microsoft accounts only
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    # Choose the right scope set for the account type.
    scopes = _PERSONAL_SCOPES if _is_personal_tenant(tenant_id) else _ORG_SCOPES

    app = msal.PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=cache,
    )

    # Try silent acquisition first (cached / refresh token).
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]

    # Fall back to device-code flow.
    flow = app.initiate_device_flow(scopes=scopes)
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
