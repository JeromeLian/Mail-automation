"""Scans mail@php-cpa.com.sg's Sent Items for Mail Collection / Mails Subscription sends, and sets the
POC for any company that doesn't have one yet, when the email's recipient matches an address already
saved on that company's profile.

Read-only against Outlook: nothing is sent, moved, or changed in the mailbox. The only write is to this
project's own presets.json / profile_emails.json / change_log.jsonl, using the same rule as the in-app
"Set POC" action (the chosen address becomes the POC; anything already saved stays on the profile).

How a match is decided, for each Sent Items email whose subject looks like a Mail Collection Reminder
or Mails Subscription send:
  1. The company name is read from the subject, e.g. "Mails Collection (Acme Pte Ltd)" -> "Acme Pte Ltd".
     This is exactly how the app fills in the subject, so this only works for emails sent from the app.
  2. That company must already be a saved profile with no POC set.
  3. The email's To address must already be saved on that company's profile (as a known extra email).
If a company matches but different sent emails used different saved addresses as the To address, it's
left for manual review (listed under "Skipped") rather than guessed.

Safe by design:
  * A dry run is the default. Nothing is written until you run with --apply.
  * Every data file is backed up before any change.
  * Requires classic Outlook open and signed in to mail@php-cpa.com.sg.

Usage:
  python scripts/backfill_poc_from_sent.py             # dry run: shows what would happen
  python scripts/backfill_poc_from_sent.py --apply     # sets the POC for every clear match
"""

import argparse
import re
import shutil
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from reminders.core import (  # noqa: E402
    CHANGE_LOG_PATH,
    PRESETS_PATH,
    PROFILE_EMAILS_PATH,
    append_change_log,
    load_json_map,
    load_presets,
    save_json_map,
    save_presets,
)
from reminders.outlook import DRAFT_ACCOUNT  # noqa: E402

OL_FOLDER_SENTMAIL = 5
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
SUBJECT_RE = re.compile(r"mails?\s+(collection|subscription)", re.IGNORECASE)
PAREN_RE = re.compile(r"\(([^()]+)\)\s*$")  # the company name the app fills in at the end of the subject


def smtp_address(recipient) -> str:
    """The recipient's real email address, even when Outlook resolved it to a cached Exchange entry."""
    try:
        addr = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        if addr:
            return addr
    except Exception:
        pass
    return recipient.Address or ""


def to_addresses(mail) -> list[str]:
    addresses = []
    for recipient in mail.Recipients:
        if recipient.Type == 1:  # olTo
            addr = smtp_address(recipient)
            if addr:
                addresses.append(addr.strip())
    return addresses


def company_from_subject(subject: str) -> str | None:
    m = PAREN_RE.search(subject)
    return m.group(1).strip() if m else None


def find_sent_items(address: str):
    import win32com.client

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    for store in ns.Stores:
        if store.DisplayName.lower() == address.lower():
            return store.GetDefaultFolder(OL_FOLDER_SENTMAIL)
    names = ", ".join(store.DisplayName for store in ns.Stores)
    raise RuntimeError(f"Outlook has no mailbox called {address}. Mailboxes found: {names}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--apply", action="store_true",
                        help="set the POC for every clear match (default is a dry run)")
    args = parser.parse_args()

    print(f"Opening Sent Items for {DRAFT_ACCOUNT}...")
    sent = find_sent_items(DRAFT_ACCOUNT)
    items = sent.Items
    items.Sort("[SentOn]", True)  # newest first, so the first match per company is the most recent one
    try:
        # Narrows down at the Outlook level before the regex check, since Sent Items can be large.
        items = items.Restrict("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%mail%'")
    except Exception:
        pass  # fall back to scanning everything if the store doesn't support this restriction

    presets = load_presets(PRESETS_PATH)
    profile_emails = load_json_map(PROFILE_EMAILS_PATH)
    groups: dict[str, list] = defaultdict(list)
    for p in presets:
        groups[p.company_name.strip().casefold()].append(p)
    no_poc_keys = {key for key, ps in groups.items() if not any(p.email.strip() for p in ps)}
    display_name = {key: ps[0].company_name for key, ps in groups.items()}

    scanned = 0
    matches: dict[str, tuple[str, str, str]] = {}  # key -> (address, subject, sent_on)
    ambiguous: dict[str, set[str]] = defaultdict(set)
    item = items.GetFirst()
    while item is not None:
        try:
            if getattr(item, "Class", None) == 43 and SUBJECT_RE.search(item.Subject or ""):
                scanned += 1
                subject = item.Subject
                company = company_from_subject(subject)
                if company:
                    key = company.strip().casefold()
                    if key in no_poc_keys:
                        known = {e.casefold() for e in profile_emails.get(key, [])}
                        for addr in to_addresses(item):
                            if addr.casefold() in known:
                                if key in matches and matches[key][0].casefold() != addr.casefold():
                                    ambiguous[key].add(matches[key][0])
                                    ambiguous[key].add(addr)
                                elif key not in matches:
                                    matches[key] = (addr, subject, str(item.SentOn))
        except Exception as e:
            print(f"  (skipped one item I couldn't read: {e})")
        item = items.GetNext()

    for key in ambiguous:
        matches.pop(key, None)

    print(f"\nScanned {scanned} Mail Collection / Mails Subscription sends.")
    print(f"Companies with no POC recorded: {len(no_poc_keys)}")
    print(f"\nClear matches ({len(matches)}):")
    for key, (addr, subject, sent_on) in matches.items():
        print(f"  - {display_name[key]}: set POC to {addr}   (from \"{subject}\", sent {sent_on})")
    if ambiguous:
        print(f"\nSkipped - more than one saved address matched, needs manual review ({len(ambiguous)}):")
        for key, addrs in ambiguous.items():
            print(f"  - {display_name[key]}: {', '.join(sorted(addrs))}")

    if not args.apply:
        print("\nDry run. Nothing was changed. Run with --apply to set these.")
        return
    if not matches:
        print("\nNothing to apply.")
        return

    for path in (PRESETS_PATH, PROFILE_EMAILS_PATH):
        if path.exists():
            shutil.copy(path, path.with_name(path.stem + ".before-sent-backfill" + path.suffix))

    for key, (addr, _subject, _sent_on) in matches.items():
        extras = profile_emails.setdefault(key, [])
        extras[:] = [e for e in extras if e.casefold() != addr.casefold()]
        for p in groups[key]:
            p.email = addr
        append_change_log("poc_set", {"company": display_name[key], "email": addr}, CHANGE_LOG_PATH)

    save_presets(PRESETS_PATH, presets)
    save_json_map(PROFILE_EMAILS_PATH, profile_emails)
    print(f"\nSet the POC for {len(matches)} compan{'y' if len(matches) == 1 else 'ies'}. "
          "Backups: *.before-sent-backfill.json")


if __name__ == "__main__":
    main()
