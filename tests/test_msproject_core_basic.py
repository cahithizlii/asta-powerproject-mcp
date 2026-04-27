"""Test core MS Project MCP infrastructure."""
import pytest
from msproject_mcp_core import _connect_app, _route_operation


def test_connect_app(msproject_app):
    """COM connection helper should return active app."""
    app = _connect_app()
    assert app is not None
    assert app.ActiveProject is not None


def test_route_operation_thresholds():
    """_route_operation should pick correct path per item count."""
    assert _route_operation(1) == "com_direct"
    assert _route_operation(5) == "com_direct"
    assert _route_operation(6) == "com_batch"
    assert _route_operation(19) == "com_batch"
    assert _route_operation(20) == "mspdi_bulk"
    assert _route_operation(500) == "mspdi_bulk"
