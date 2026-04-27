"""Pytest fixtures shared across MS Project tests."""
import pytest
import os

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@pytest.fixture(scope="session")
def fixtures_dir() -> str:
    return os.path.join(REPO_ROOT, "tests", "fixtures")


@pytest.fixture(scope="session")
def msproject_app():
    """Session-level MS Project COM connection. Skips if not available."""
    pythoncom = None
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            app = win32com.client.GetActiveObject("MSProject.Application")
        except Exception as com_err:
            # COM-specific failure (MS Project not running, etc.)
            pytest.skip(f"MS Project not available: {com_err}")
        if app.ActiveProject is None:
            pytest.skip("No active MS Project project")
        yield app
    except ImportError as e:
        pytest.skip(f"COM bindings unavailable: {e}")
    finally:
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
