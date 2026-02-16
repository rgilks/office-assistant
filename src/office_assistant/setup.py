"""Interactive setup for Office Assistant authentication.

Run with: uv run python -m office_assistant.setup
"""

from __future__ import annotations

import sys
from pathlib import Path

from office_assistant.auth import (
    CACHE_FILE,
    AuthenticationRequired,
    complete_device_flow,
    get_token,
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
    return config.get("CLIENT_ID") or "", config.get("TENANT_ID") or ""


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
        print("Your credentials file is incomplete. Let's fill in the missing values.")
    else:
        print("Let's connect your Microsoft account.")
        print()
        print("You'll need an Application (client) ID from an Azure App Registration.")
        print("If you haven't created one yet, see the README or type /calendar-setup")
        print("in Claude Code for step-by-step instructions.")
    print()

    client_id = existing_client_id
    if not client_id:
        print("You can find this in the Azure Portal under App registrations")
        print("> your app > Overview.")
        print()
        client_id = input("Application (client) ID: ").strip()
        if not client_id:
            print("Error: A client ID is required to continue.", file=sys.stderr)
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
    """Ensure we have a valid token, prompting the user if needed."""
    try:
        get_token()
        print()
        print("Already signed in.")
    except AuthenticationRequired as auth_req:
        print()
        if CACHE_FILE.exists():
            print("Your sign-in has expired. Let's reconnect.")
        print()
        print(auth_req.message)
        print()
        print("Waiting for you to complete sign-in in your browser...")
        complete_device_flow(auth_req.flow)
        print()
        print("Signed in successfully!")


def main() -> None:
    if not _env_is_complete():
        _create_env_file()

    _authenticate()

    print()
    print("You're all set! Start a new Claude Code conversation and type")
    print("/calendar to manage your calendar.")


if __name__ == "__main__":
    main()
