"""Tests for shared helper functions."""

from __future__ import annotations

from office_assistant.graph_client import GraphApiError
from office_assistant.tools._helpers import (
    _coerce_datetime,
    _parse_iso_datetime,
    graph_error_response,
    validate_datetime,
    validate_datetime_order,
    validate_timezone,
)


class TestParseIsoDatetime:
    def test_z_suffix_converted_to_utc(self):
        dt = _parse_iso_datetime("2026-02-16T09:00:00Z")
        assert dt.tzinfo is not None
        assert dt.isoformat() == "2026-02-16T09:00:00+00:00"

    def test_naive_datetime(self):
        dt = _parse_iso_datetime("2026-02-16T09:00:00")
        assert dt.tzinfo is None

    def test_offset_datetime(self):
        dt = _parse_iso_datetime("2026-02-16T09:00:00+05:30")
        assert dt.tzinfo is not None


class TestValidateDatetime:
    def test_valid_datetime(self):
        assert validate_datetime("2026-02-16T09:00:00", "start") is None

    def test_invalid_datetime(self):
        err = validate_datetime("not-a-date", "start")
        assert err is not None
        assert "ISO 8601" in err


class TestValidateTimezone:
    def test_valid_timezone(self):
        assert validate_timezone("Europe/London", "tz") is None

    def test_invalid_timezone(self):
        err = validate_timezone("Not/AZone", "tz")
        assert err is not None
        assert "IANA timezone" in err


class TestCoerceDatetime:
    def test_no_timezone_returns_as_is(self):
        dt = _parse_iso_datetime("2026-02-16T09:00:00")
        result = _coerce_datetime(dt, None)
        assert result.tzinfo is None

    def test_naive_datetime_gets_timezone(self):
        dt = _parse_iso_datetime("2026-02-16T09:00:00")
        result = _coerce_datetime(dt, "Europe/London")
        assert result.tzinfo is not None


class TestValidateDatetimeOrder:
    def test_invalid_start_datetime(self):
        err = validate_datetime_order("bad", "2026-02-16T09:00:00")
        assert err is not None
        assert "ISO 8601" in err

    def test_invalid_end_datetime(self):
        err = validate_datetime_order("2026-02-16T09:00:00", "bad")
        assert err is not None
        assert "ISO 8601" in err

    def test_invalid_start_timezone(self):
        err = validate_datetime_order(
            "2026-02-16T09:00:00",
            "2026-02-16T10:00:00",
            start_timezone="Not/AZone",
        )
        assert err is not None
        assert "IANA timezone" in err

    def test_invalid_end_timezone(self):
        err = validate_datetime_order(
            "2026-02-16T09:00:00",
            "2026-02-16T10:00:00",
            end_timezone="Not/AZone",
        )
        assert err is not None
        assert "IANA timezone" in err

    def test_mismatched_timezone_awareness(self):
        err = validate_datetime_order(
            "2026-02-16T09:00:00+00:00",
            "2026-02-16T10:00:00",
        )
        assert err is not None
        assert "timezone offsets" in err


class TestGraphErrorResponse:
    def test_validation_error_type(self):
        exc = GraphApiError(
            status_code=400,
            message="Invalid parameter",
            code="ErrorInvalidRequest",
        )
        result = graph_error_response(exc)
        assert result["errorType"] == "validation_error"

    def test_generic_error_type(self):
        exc = GraphApiError(
            status_code=502,
            message="Bad gateway",
            code="BadGateway",
        )
        result = graph_error_response(exc)
        assert result["errorType"] == "graph_error"

    def test_throttled_error_type(self):
        exc = GraphApiError(
            status_code=429,
            message="Too many requests",
            code="TooManyRequests",
            retry_after_seconds=30,
        )
        result = graph_error_response(exc)
        assert result["errorType"] == "throttled"
        assert result["retryAfterSeconds"] == 30

    def test_not_found_error_type(self):
        exc = GraphApiError(
            status_code=404,
            message="Not found",
            code="ErrorItemNotFound",
        )
        result = graph_error_response(exc)
        assert result["errorType"] == "not_found"

    def test_fallback_message_used(self):
        exc = GraphApiError(status_code=403, message="Forbidden")
        result = graph_error_response(exc, fallback_message="Custom message")
        assert result["error"] == "Custom message"

    def test_error_code_included(self):
        exc = GraphApiError(status_code=500, message="Error", code="SomeCode")
        result = graph_error_response(exc)
        assert result["errorCode"] == "SomeCode"

    def test_request_id_included(self):
        exc = GraphApiError(status_code=500, message="Error", request_id="req-42")
        result = graph_error_response(exc)
        assert result["requestId"] == "req-42"
