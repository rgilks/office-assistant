"""Tests for the authentication module."""

from __future__ import annotations

import json
from unittest.mock import MagicMock, patch

import pytest

from office_assistant.auth import (
    AuthenticationRequired,
    _build_cache,
    _load_env,
    _save_cache,
    clear_cache,
    get_token,
)


class TestTokenCache:
    def test_build_cache_no_file(self, tmp_path):
        """Cache is empty when no file exists."""
        with patch("office_assistant.auth.CACHE_FILE", tmp_path / "nope.json"):
            cache = _build_cache()
            assert cache.serialize() == "{}"

    def test_build_cache_from_file(self, tmp_path):
        cache_file = tmp_path / "token_cache.json"
        cache_file.write_text('{"some": "data"}')
        with patch("office_assistant.auth.CACHE_FILE", cache_file):
            cache = _build_cache()
            assert cache is not None

    def test_build_cache_corrupt_file(self, tmp_path):
        cache_file = tmp_path / "token_cache.json"
        cache_file.write_text("not-json")
        with patch("office_assistant.auth.CACHE_FILE", cache_file):
            cache = _build_cache()
            assert cache.serialize() == "{}"

    def test_save_cache_creates_dir(self, tmp_path):
        cache_dir = tmp_path / "new_dir"
        cache_file = cache_dir / "token_cache.json"
        mock_cache = MagicMock()
        mock_cache.has_state_changed = True
        mock_cache.serialize.return_value = '{"token": "data"}'

        with (
            patch("office_assistant.auth.CACHE_DIR", cache_dir),
            patch("office_assistant.auth.CACHE_FILE", cache_file),
        ):
            _save_cache(mock_cache)

        assert cache_dir.exists()
        assert json.loads(cache_file.read_text()) == {"token": "data"}

    def test_save_cache_skips_when_no_change(self, tmp_path):
        cache_file = tmp_path / "token_cache.json"
        mock_cache = MagicMock()
        mock_cache.has_state_changed = False

        with patch("office_assistant.auth.CACHE_FILE", cache_file):
            _save_cache(mock_cache)

        assert not cache_file.exists()

    def test_clear_cache(self, tmp_path):
        cache_file = tmp_path / "token_cache.json"
        cache_file.write_text("{}")
        with patch("office_assistant.auth.CACHE_FILE", cache_file):
            assert clear_cache() is True
            assert not cache_file.exists()

    def test_clear_cache_no_file(self, tmp_path):
        with patch("office_assistant.auth.CACHE_FILE", tmp_path / "nope.json"):
            assert clear_cache() is False


class TestLoadEnv:
    def test_missing_env_raises(self, tmp_path):
        env_file = tmp_path / ".env"
        env_file.write_text("")
        with (
            patch.dict("os.environ", {"DOTENV_PATH": str(env_file)}),
            pytest.raises(RuntimeError, match="Microsoft account isn't configured"),
        ):
            _load_env()

    def test_loads_credentials(self, tmp_path):
        env_file = tmp_path / ".env"
        env_file.write_text("CLIENT_ID=abc\nTENANT_ID=xyz\n")
        with patch.dict("os.environ", {"DOTENV_PATH": str(env_file)}):
            client_id, tenant_id = _load_env()
            assert client_id == "abc"
            assert tenant_id == "xyz"


class TestGetToken:
    def test_silent_token_acquisition(self, patch_auth):
        """When a cached token exists, acquire_token_silent succeeds."""
        mock_app = MagicMock()
        mock_app.get_accounts.return_value = [{"username": "user@test.com"}]
        mock_app.acquire_token_silent.return_value = {"access_token": "cached-token"}

        with (
            patch("office_assistant.auth._load_env", return_value=("client-id", "tenant-id")),
            patch("office_assistant.auth._build_cache"),
            patch("office_assistant.auth._save_cache"),
            patch("msal.PublicClientApplication", return_value=mock_app),
        ):
            # Stop the autouse fixture and call the real get_token
            patch_auth.stop()
            try:
                token = get_token()
                assert token == "cached-token"
            finally:
                patch_auth.start()

    def test_device_code_flow_raises_auth_required(self, patch_auth):
        """When no cached token, raises AuthenticationRequired with sign-in details."""
        mock_app = MagicMock()
        mock_app.get_accounts.return_value = []
        mock_app.initiate_device_flow.return_value = {
            "user_code": "ABC123",
            "verification_uri": "https://microsoft.com/devicelogin",
            "message": "Go to https://microsoft.com/devicelogin and enter ABC123",
        }

        with (
            patch("office_assistant.auth._load_env", return_value=("client-id", "tenant-id")),
            patch("office_assistant.auth._build_cache"),
            patch("msal.PublicClientApplication", return_value=mock_app),
        ):
            patch_auth.stop()
            try:
                with pytest.raises(AuthenticationRequired) as exc_info:
                    get_token()
                assert exc_info.value.user_code == "ABC123"
                assert exc_info.value.url == "https://microsoft.com/devicelogin"
                assert exc_info.value.flow is not None
                mock_app.initiate_device_flow.assert_called_once()
            finally:
                patch_auth.start()

    def test_device_code_flow_initiation_failure(self, patch_auth):
        """Raises RuntimeError when device code flow can't be started."""
        mock_app = MagicMock()
        mock_app.get_accounts.return_value = []
        mock_app.initiate_device_flow.return_value = {
            "error_description": "Application is not configured"
        }

        with (
            patch("office_assistant.auth._load_env", return_value=("client-id", "tenant-id")),
            patch("office_assistant.auth._build_cache"),
            patch("msal.PublicClientApplication", return_value=mock_app),
        ):
            patch_auth.stop()
            try:
                with pytest.raises(RuntimeError, match="Application is not configured"):
                    get_token()
            finally:
                patch_auth.start()
