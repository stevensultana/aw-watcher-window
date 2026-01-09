import pywintypes
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")

def get_outlook_activity() -> dict:
    global outlook
    try:
        explorer = outlook.ActiveExplorer()
    except pywintypes.com_error as e:
        if e.hresult == -2147023174:  # The RPC server is unavailable
            outlook = win32com.client.Dispatch("Outlook.Application")
            explorer = outlook.ActiveExplorer()
        else:
            raise e

    if explorer is None:
        return {}

    selection = explorer.Selection
    if selection.Count != 1:
        return {}

    item = selection.Item(1)

    return {
        "title": item.Subject,
        "folder": explorer.CurrentFolder.Name,
    }
