"""Pytest fixtures shared across MS Project tests."""
import pytest
import os

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@pytest.fixture(scope="session")
def fixtures_dir():
    return os.path.join(REPO_ROOT, "tests", "fixtures")


@pytest.fixture(scope="session")
def msproject_app():
    """Session-level MS Project COM connection. Skips if not available."""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.GetActiveObject("MSProject.Application")
        if app.ActiveProject is None:
            pytest.skip("No active MS Project")
        return app
    except Exception as e:
        pytest.skip(f"MS Project not available: {e}")
