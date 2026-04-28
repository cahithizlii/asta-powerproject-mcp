"""Unit test for _format_com_error helper (no MS Project required)."""
import pytest
from msproject_mcp_core import _format_com_error


def test_format_pywintypes_com_error_tuple():
    """pywintypes.com_error has args = (hresult, msg, excepinfo, argerr).
    excepinfo[2] is the human-readable description."""
    class FakeComError(Exception):
        def __init__(self):
            self.args = (
                -2147352567,
                "Exception occurred.",
                (0, "Microsoft Project", "The calendar name already exists.",
                 None, 0, -2147352567),
                None,
            )
    result = _format_com_error(FakeComError())
    assert result == "The calendar name already exists."


def test_format_falls_back_to_args1_when_excepinfo_empty_description():
    """If excepinfo[2] (description) is None/empty, fall back to args[1]."""
    class FakeComError(Exception):
        def __init__(self):
            self.args = (
                -2147352567,
                "Exception occurred.",
                (0, "Microsoft Project", None, None, 0, -2147352567),
                None,
            )
    result = _format_com_error(FakeComError())
    assert result == "Exception occurred."


def test_format_plain_exception_uses_str():
    """Plain Python exception falls through to str(e)."""
    e = ValueError("bad input")
    assert _format_com_error(e) == "bad input"


def test_format_empty_args_does_not_crash():
    """Edge: exception with no args."""
    class Empty(Exception):
        pass
    e = Empty()
    result = _format_com_error(e)
    assert isinstance(result, str)


def test_format_args_too_short_for_excepinfo():
    """args has only 1 element, no excepinfo tuple."""
    class Short(Exception):
        def __init__(self):
            self.args = ("just one message",)
    result = _format_com_error(Short())
    assert "just one message" in result
