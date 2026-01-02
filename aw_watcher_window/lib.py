import sys
from typing import Optional

from .exceptions import FatalError


def get_current_window_linux() -> Optional[dict]:
    from . import xlib

    window = xlib.get_current_window()

    if window is None:
        cls = "unknown"
        name = "unknown"
    else:
        cls = xlib.get_window_class(window)
        name = xlib.get_window_name(window)

    return {"app": cls, "title": name}


def get_current_window_macos(strategy: str) -> Optional[dict]:
    # TODO should we use unknown when the title is blank like the other platforms?

    # `jxa` is the default & preferred strategy. It includes the url + incognito status
    if strategy == "jxa":
        from . import macos_jxa

        return macos_jxa.getInfo()
    elif strategy == "applescript":
        from . import macos_applescript

        return macos_applescript.getInfo()
    else:
        raise FatalError(f"invalid strategy '{strategy}'")


def get_current_window_windows() -> Optional[dict]:
    from . import windows
    import pywintypes
    import wmi

    window_handle = windows.get_active_window_handle()
    if window_handle == 0:
        return None
        # c = wmi.WMI()
        # for process in c.Win32_Process(name="consent.exe"):
        #     return {"app": "consent.exe", "title": "UAC"}
        # for process in c.Win32_Process(name="logonui.exe"):
        #     return {"app": "logonUI.exe", "title": "Logon Screen"}
        # return {"app": "unknown", "title": "unknown"}

    try:
        app = windows.get_app_name(window_handle)
    except pywintypes.error as e:
        # try with wmi method
        if e.winerror == 5:  # Access is denied
            app = windows.get_app_name_wmi(window_handle)
        else:
            raise e

    title = windows.get_window_title(window_handle)

    if app is None:
        app = "unknown"
    if title is None:
        title = "unknown"

    return {"app": app, "title": title}


def get_current_window(strategy: Optional[str] = None) -> Optional[dict]:
    """
    :raises FatalError: if a fatal error occurs (e.g. unsupported platform, X server closed)
    """

    if sys.platform.startswith("linux"):
        return get_current_window_linux()
    elif sys.platform == "darwin":
        if strategy is None:
            raise FatalError("macOS strategy not specified")
        return get_current_window_macos(strategy)
    elif sys.platform in ["win32", "cygwin"]:
        return get_current_window_windows()
    else:
        raise FatalError(f"Unknown platform: {sys.platform}")
