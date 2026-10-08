"""Applies a salesperson's exported Change Log to the central data in this project folder.

Each person's installed copy keeps its own data and logs every change it makes (new or edited presets,
POC changes, profile emails, person names, mail subscriptions, template and profile deletions) to
change_log.jsonl. Export Change Log in the app, from the Change Log button on the start screen, and run
this script against the exported file.

The app itself can also open an exported log directly, from the Change Log window, and review each
change one at a time or in a batch before approving it. This script is the same thing from the command
line, for applying everything at once.

What it can apply automatically: preset changes, POC changes, profile emails, person names, mail
subscription status, and profile or template deletions. These are stored as plain data, so replaying
them here is exact.

What it cannot apply: a template's subject, body or attachment requirement. Only the fact that a
template changed is logged, not its content, since the full content can include pasted images that
only exist on that person's PC. These are always listed for you to apply by hand, by copying the
template in the app.

Safe by design:
  * A dry run is the default. Nothing is written until you run with --apply.
  * Every data file is backed up before any change.
  * Entries are applied in the order they happened. Later entries for the same thing naturally win,
    the same as if the change had been made here directly.

Usage:
  python scripts/apply_change_log.py change_log.jsonl            # dry run: shows what would happen
  python scripts/apply_change_log.py change_log.jsonl --apply    # applies the safe entries
"""

import argparse
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from reminders.core import (  # noqa: E402
    CHANGE_LOG_MANUAL_ACTIONS,
    PERSON_NAMES_PATH,
    PRESETS_PATH,
    PROFILE_EMAILS_PATH,
    SUBSCRIPTIONS_PATH,
    apply_change_entry,
    load_change_log,
    load_json_map,
    load_presets,
    save_json_map,
    save_presets,
)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("log_file", help="the change_log.jsonl exported from a salesperson's install")
    parser.add_argument("--apply", action="store_true", help="make the change (default is a dry run)")
    args = parser.parse_args()

    log_path = Path(args.log_file)
    if not log_path.exists():
        sys.exit(f"File not found: {log_path}")
    entries = sorted(load_change_log(log_path), key=lambda e: e.get("timestamp", ""))
    manual = [e for e in entries if e.get("action") in CHANGE_LOG_MANUAL_ACTIONS]
    automatic = [e for e in entries if e.get("action") not in CHANGE_LOG_MANUAL_ACTIONS]

    presets = load_presets(PRESETS_PATH)
    profile_emails = load_json_map(PROFILE_EMAILS_PATH)
    person_names = load_json_map(PERSON_NAMES_PATH)
    subscriptions = load_json_map(SUBSCRIPTIONS_PATH)

    applied = skipped = 0
    skip_notes: list[str] = []
    for entry in automatic:
        ok, note = apply_change_entry(entry, presets, profile_emails, person_names, subscriptions)
        if ok:
            applied += 1
        else:
            skipped += 1
            skip_notes.append(f"  - {entry.get('timestamp', '')}  {entry.get('action', '')} "
                              f"({entry.get('company', entry.get('email', ''))}): {note}")

    print(f"Entries in the log:                 {len(entries)}")
    print(f"Applied automatically:              {applied}")
    print(f"Skipped (already matched the data, or not found): {skipped}")
    for note in skip_notes[:20]:
        print(note)
    if len(skip_notes) > 20:
        print(f"  ... and {len(skip_notes) - 20} more")

    print(f"\nNeeds manual review (template content, not in the log): {len(manual)}")
    for entry in manual:
        action = entry.get("action", "")
        if action == "template_saved":
            print(f"  - Template '{entry.get('name', '')}' was {entry.get('action_type', 'changed')} "
                  "on their PC. Copy its subject, body and settings into this template yourself.")
        elif action == "template_backfilled":
            print(f"  - Template '{entry.get('template', '')}' was added to "
                  f"{entry.get('companies_added', 0)} companies on their PC. Check it exists here, "
                  "and use Create template > add new template's presets if not.")
        elif action == "preset_template_renamed":
            print(f"  - Template '{entry.get('old', '')}' was renamed to '{entry.get('new', '')}' "
                  "on their PC. Rename it here the same way if that's correct.")

    if not args.apply:
        print("\nDry run. Nothing was changed. Run with --apply to make the automatic changes.")
        return
    if applied == 0:
        print("\nNothing to apply.")
        return

    for path in (PRESETS_PATH, PROFILE_EMAILS_PATH, PERSON_NAMES_PATH, SUBSCRIPTIONS_PATH):
        if path.exists():
            shutil.copy(path, path.with_name(path.stem + ".before-change-log" + path.suffix))
    save_presets(PRESETS_PATH, presets)
    save_json_map(PROFILE_EMAILS_PATH, profile_emails)
    save_json_map(PERSON_NAMES_PATH, person_names)
    save_json_map(SUBSCRIPTIONS_PATH, subscriptions)
    print(f"\nApplied {applied} change(s). Backups: *.before-change-log.json")


if __name__ == "__main__":
    main()
