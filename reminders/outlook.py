"""Sends or opens a mail through the Outlook desktop app on this PC.

Only classic Outlook is supported (it uses the Outlook COM interface).
"""

from pathlib import Path

OL_MAIL_ITEM = 0
OL_FOLDER_DRAFTS = 16
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

# Drafts are saved to this mailbox's Drafts folder. It must be an account added to classic Outlook.
DRAFT_ACCOUNT = "mail@acme-cpa.com.sg"


def _find_draft_folder(outlook, address: str):
    """The Drafts folder of the mailbox called `address`. Raises if Outlook has no such mailbox."""
    ns = outlook.GetNamespace("MAPI")
    for store in ns.Stores:
        if store.DisplayName.lower() == address.lower():
            return store.GetDefaultFolder(OL_FOLDER_DRAFTS)
    names = ", ".join(store.DisplayName for store in ns.Stores)
    raise RuntimeError(f"Outlook has no mailbox called {address}. Mailboxes found: {names}")


def _sending_account(outlook, address: str):
    for account in outlook.Session.Accounts:
        if account.SmtpAddress.lower() == address.lower():
            return account
    return None


def save_draft(to: list[str], cc: list[str], subject: str, html_body: str,
               images: list[tuple[str, Path]], attachments: list[str]) -> None:
    """Saves the mail to the Drafts folder of DRAFT_ACCOUNT. Nothing is sent.

    images: (content id, file) pairs. Each is embedded in the body where cid:<content id> appears.
    attachments: file paths added as normal attachments.
    """
    # Imported here so the rest of the app still runs (and can be tested) without pywin32.
    import win32com.client

    outlook = win32com.client.Dispatch("Outlook.Application")
    drafts = _find_draft_folder(outlook, DRAFT_ACCOUNT)
    mail = drafts.Items.Add(OL_MAIL_ITEM)
    account = _sending_account(outlook, DRAFT_ACCOUNT)
    if account is not None:
        mail.SendUsingAccount = account  # so the draft is set to send from this account
    mail.To = "; ".join(to)
    mail.CC = "; ".join(cc)
    mail.Subject = subject
    for cid, path in images:
        attachment = mail.Attachments.Add(str(path))
        attachment.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)
        attachment.PropertyAccessor.SetProperty(PR_ATTACHMENT_HIDDEN, True)  # inline, not in the attachment list
    for path in attachments:
        mail.Attachments.Add(str(path))
    mail.HTMLBody = html_body
    mail.Save()
