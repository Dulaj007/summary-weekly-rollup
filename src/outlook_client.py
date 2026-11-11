"""Thin wrapper around the running Outlook desktop app via COM."""
import win32com.client

OL_FOLDER_INBOX = 6


def connect():
    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")


def get_folder(namespace, folder_path: str):
    """Resolve a folder path like 'Inbox', 'Inbox/Reports', or
    'sharedmailbox@company.com/Inbox/Reports' to an Outlook folder object."""
    parts = [p for p in folder_path.split("/") if p]
    if not parts:
        raise ValueError("folder_path must not be empty")

    top_level_names = {f.Name.lower(): f for f in namespace.Folders}
    if parts[0].lower() in top_level_names:
        folder = top_level_names[parts[0].lower()]
        remaining = parts[1:]
    else:
        folder = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
        remaining = parts
        if remaining and remaining[0].lower() == "inbox":
            remaining = remaining[1:]

    for part in remaining:
        folder = folder.Folders[part]
    return folder


def get_items_in_range(folder, start_dt, end_dt):
    """Return MailItems received within [start_dt, end_dt), sorted ascending."""
    items = folder.Items
    items.Sort("[ReceivedTime]")
    fmt = "%m/%d/%Y %I:%M %p"
    restriction = (
        f"[ReceivedTime] >= '{start_dt.strftime(fmt)}' "
        f"AND [ReceivedTime] < '{end_dt.strftime(fmt)}'"
    )
    return list(items.Restrict(restriction))
