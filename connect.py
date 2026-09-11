import subprocess
import time
import win32com.client

SAP_LOGON_PATH = (
    r"C:\Program Files (x86)\SAP\FrontEnd"
    r"\SapGui\saplogon.exe"
)

SAP_CONNECTION_NAME = "M04 [SSO]"


class SAPConnectionError(Exception):
    pass


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


def get_session():
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


