import subprocess
import sys
import time
import win32com.client

SAP_LOGON_PATH = (
    r"C:\Program Files (x86)\SAP\FrontEnd"
    r"\SapGui\saplogon.exe"
)

SAP_CONNECTION_NAME = "M04 [SSO]"

# Slow mode: seconds to wait after every SAP action (press, Enter, typing, selecting)
STEP_DELAY = 0.5

class SAPConnectionError(Exception):
    pass


class SlowSAP:
    """Wraps a SAP GUI object: waits STEP_DELAY after actions, reads are not delayed."""
    ACTIONS = {"press", "sendVKey", "setFocus", "select"}
    SET_ACTIONS = {"text", "selected", "Position", "caretPosition", "key"}

    def __init__(self, obj):
        object.__setattr__(self, "_obj", obj)

    def __getattr__(self, name):
        value = getattr(self._obj, name)
        # COM objects (Info, VerticalScrollbar, ...) are callable too, so check them first
        if hasattr(value, "_oleobj_"):
            return SlowSAP(value)
        if callable(value):
            def call(*args):
                result = value(*args)
                if name in self.ACTIONS:
                    print(f"  [slow] {name}{args}")
                    time.sleep(STEP_DELAY)
                return _wrap(result)
            return call
        return value

    def __setattr__(self, name, value):
        setattr(self._obj, name, value)
        if name in self.SET_ACTIONS:
            print(f"  [slow] {name} = {value!r}")
            time.sleep(STEP_DELAY)

    def __iter__(self):
        # Collections (session.Children, usr.Children): Python looks up __iter__ on the class
        for item in self._obj:
            yield _wrap(item)


def _wrap(value):
    # Wrap SAP COM objects (findById results, rows, scrollbars); leave str/int/None as is
    if hasattr(value, "_oleobj_"):
        return SlowSAP(value)
    return value


def slow_from_args():
    """True if the script was started with 'slow', e.g. python routing_main.py slow"""
    return "slow" in sys.argv


def _is_sap_logon_running():
    """Check if saplogon.exe is already running."""
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq saplogon.exe"],
        capture_output=True,
        text=True
    )
    return "saplogon.exe" in result.stdout.lower()


def _launch_sap_logon():
    """Launch SAP Logon if not already running."""
    if _is_sap_logon_running():
        return

    subprocess.Popen([SAP_LOGON_PATH])
    # Wait for SAP Logon to initialize
    time.sleep(5)


def _get_sap_application(retries=5, delay=2):
    """Get the SAP GUI scripting engine, retrying if needed."""
    for i in range(retries):
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            if sap_gui:
                return sap_gui.GetScriptingEngine
        except Exception:
            if i < retries - 1:
                time.sleep(delay)
    raise SAPConnectionError(
        "Could not connect to SAP GUI scripting engine"
    )


def _open_connection(application):
    """Open the M04 SSO connection from SAP Logon."""
    connection = application.OpenConnection(
        SAP_CONNECTION_NAME, True
    )
    # Wait for the session to be ready
    time.sleep(3)
    return connection


def get_session(slow=False):
    """SAP session. slow=True waits STEP_DELAY seconds after every action."""
    session = _get_session()
    if slow:
        print(f"Slow mode: {STEP_DELAY}s after every SAP action")
        return SlowSAP(session)
    return session


def _get_session():
    try:
        # Launch SAP Logon if not running
        _launch_sap_logon()

        # Get scripting engine
        application = _get_sap_application()

        # Check if there's already an active connection
        if application.Children.Count > 0:
            connection = application.Children(0)
            if connection.Children.Count > 0:
                return connection.Children(0)

        # No active session - open M04 connection
        connection = _open_connection(application)

        if connection.Children.Count == 0:
            raise SAPConnectionError(
                "Connection opened but no session available"
            )

        return connection.Children(0)

    except SAPConnectionError:
        raise

    except Exception as e:
        raise SAPConnectionError(
            f"Failed to connect to SAP: {e}"
        )


