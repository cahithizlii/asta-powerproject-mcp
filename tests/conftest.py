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


@pytest.fixture
def clean_test_project(msproject_app):
    """Function-scoped fixture providing an ISOLATED empty MS Project for tests.

    SAFETY: Creates a brand new project via FileNew, leaves user's projects untouched.
    On teardown: closes test project without saving, restores user's original active
    project. Tests using this fixture will NEVER modify the user's real work.
    """
    app = msproject_app
    # Remember user's active project so we can restore focus on teardown
    original_name = None
    try:
        if app.ActiveProject is not None:
            original_name = app.ActiveProject.Name
    except Exception:
        pass

    # Create a fresh isolated project (typically named "ProjectN")
    app.FileNew()
    test_proj = app.ActiveProject
    test_name = test_proj.Name

    yield test_proj

    # Teardown: close test project without saving, restore user's original
    try:
        # Find and activate the test project window, then close without save
        for i in range(1, app.Projects.Count + 1):
            try:
                if app.Projects(i).Name == test_name:
                    try:
                        app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    except Exception:
                        pass
                    try:
                        app.FileClose(0)  # 0 = pjDoNotSave
                    except Exception:
                        pass
                    break
            except Exception:
                continue
        # Restore user's original window focus
        if original_name:
            for i in range(1, app.Projects.Count + 1):
                try:
                    if app.Projects(i).Name == original_name:
                        try:
                            app.WindowActivate(app.Projects(i).Windows(1).Caption)
                        except Exception:
                            pass
                        break
                except Exception:
                    continue
    except Exception:
        pass  # Never let teardown fail tests
