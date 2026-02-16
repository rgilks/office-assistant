"""Interactive setup for Office Assistant authentication.

Run with: uv run python -m office_assistant.setup
"""

from __future__ import annotations

import sys
from pathlib import Path

from office_assistant.auth import (
    CACHE_FILE,
    _build_cache,
    _is_personal_tenant,
    _load_env,
    _save_cache,
)


def _dotenv_path() -> Path:
    """Resolve the .env file path."""
    import os

    return Path(os.environ.get("DOTENV_PATH", ".env"))


def _load_existing_env() -> tuple[str, str]:
    """Load CLIENT_ID and TENANT_ID from existing .env file."""
    path = _dotenv_path()
    if not path.exists():
        return "", ""
    from dotenv import dotenv_values

    config = dotenv_values(path)
    return config.get("CLIENT_ID", ""), config.get("TENANT_ID", "")


def _env_is_complete() -> bool:
    client_id, tenant_id = _load_existing_env()
    return bool(client_id) and bool(tenant_id)


def _create_env_file() -> None:
    """Prompt the user for credentials and write the .env file."""
    existing_client_id, existing_tenant_id = _load_existing_env()

    if existing_client_id and existing_tenant_id:
        return  # Already complete

    print()
    if existing_client_id or existing_tenant_id:
        print("Your .env file is incomplete. Let's fill in the missing values.")
    else:
        print("No .env file found. Let's set one up now.")
    print()
    print("You'll need an Azure App Registration first.")
    print("See /calendar-setup in Claude Code for step-by-step instructions,")
    print("or visit: https://portal.azure.com → App registrations")
    print()

    client_id = existing_client_id
    if not client_id:
        client_id = input("Paste your Application (client) ID: ").strip()
        if not client_id:
            print("Error: CLIENT_ID is required.", file=sys.stderr)
            sys.exit(1)

    tenant_id = existing_tenant_id
    if not tenant_id:
        print()
        print("Account type:")
        print("  1. Work/school account (Microsoft 365)")
        print("  2. Personal account (outlook.com, hotmail.com, live.com)")
        print()
        choice = input("Enter 1 or 2: ").strip()

        if choice == "2":
            tenant_id = "consumers"
        else:
            tenant_id = input("Paste your Directory (tenant) ID: ").strip()
            if not tenant_id:
                print("Error: TENANT_ID is required.", file=sys.stderr)
                sys.exit(1)

    path = _dotenv_path()
    path.write_text(
        f"# Azure App Registration credentials\n"
        f"# Run /calendar-setup for full instructions\n"
        f"\n"
        f"CLIENT_ID={client_id}\n"
        f"TENANT_ID={tenant_id}\n"
    )
    path.chmod(0o600)
    print()
    print(f"Saved credentials to {path}")


def _authenticate() -> None:
    """Run the device code flow and cache the token."""
    import msal

    client_id, tenant_id = _load_env()
    cache = _build_cache()

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    if _is_personal_tenant(tenant_id):
        scopes = ["Calendars.ReadWrite", "User.Read"]
    else:
        scopes = [
            "Calendars.ReadWrite",
            "Calendars.ReadWrite.Shared",
            "Place.Read.All",
            "User.Read",
        ]

    app = msal.PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=cache,
    )

    # Check if already authenticated.
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            print()
            print(f"Already authenticated (token cached at {CACHE_FILE}).")
            return

    # Start device code flow.
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        print(
            f"Error: {flow.get('error_description', 'Could not start device code flow.')}",
            file=sys.stderr,
        )
        sys.exit(1)

    print()
    print(flow["message"])
    print()
    print("Waiting for you to sign in...")

    result = app.acquire_token_by_device_flow(flow)
    _save_cache(cache)

    if "access_token" in result:
        print()
        print("Authenticated successfully!")
        print(f"Token cached at {CACHE_FILE}")
    else:
        print(
            f"Error: {result.get('error_description', 'Authentication failed.')}",
            file=sys.stderr,
        )
        sys.exit(1)


def main() -> None:
    print("=== Office Assistant Setup ===")

    if not _env_is_complete():
        _create_env_file()

    _authenticate()

    print()
    print("You're all set! Use /calendar in Claude Code to get started.")


if __name__ == "__main__":
    main()
