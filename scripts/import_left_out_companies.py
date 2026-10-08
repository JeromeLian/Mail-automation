"""Add the companies left out of the Zoho list as profiles with no POC.

Safe by design:
  * Companies that already have a profile are never changed. This includes near-matches on name
    (for example the same name with or without "Pte Ltd"), which are listed for you to review.
  * New profiles get a preset for every template with no POC, so the app shows the "no POC" alert.
  * The Zoho emails are saved as choices for the POC picker. @php-cpa.com.sg addresses are never used.
  * A dry run is the default. Nothing is written until you run with --apply.
  * presets.json and profile_emails.json are backed up before any change.

Usage:
  python scripts/import_left_out_companies.py            # dry run: shows what would happen
  python scripts/import_left_out_companies.py --apply    # makes the change
"""

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from reminders.core import (  # noqa: E402
    PRESETS_PATH,
    PROFILE_EMAILS_PATH,
    Preset,
    load_presets,
    load_profile_emails,
    load_templates,
    save_presets,
    save_profile_emails,
)

LEFT_OUT_PATH = ROOT / "left_out_companies.json"
FIRM_DOMAIN = "@php-cpa.com.sg"
SUFFIX_WORDS = {"pte", "ltd", "limited", "sdn", "bhd", "private", "the"}
# Test and placeholder records in the Zoho data. These are left out of the import.
TEST_NAME = re.compile(r"\btest\b|^a{3,}|^abc|dummy|sample|admin|^asdf|^xxx|^zzz", re.IGNORECASE)


def normal(name: str) -> str:
    """Name with punctuation, capitals and company suffixes removed, for spotting near-duplicates."""
    words = re.sub(r"[^a-z0-9 ]", " ", name.casefold()).split()
    return " ".join(w for w in words if w not in SUFFIX_WORDS)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--apply", action="store_true", help="make the change (default is a dry run)")
    args = parser.parse_args()

    presets = load_presets(PRESETS_PATH)
    templates = [t.name for t in load_templates(ROOT / "templates.json")]
    profile_emails = load_profile_emails(PROFILE_EMAILS_PATH)
    left_out = json.loads(LEFT_OUT_PATH.read_text(encoding="utf-8"))

    existing_keys = {p.company_name.strip().casefold() for p in presets}
    existing_normal = {normal(p.company_name) for p in presets}

    to_add: list[dict] = []
    skipped_existing: list[str] = []
    skipped_near: list[str] = []
    skipped_no_email: list[str] = []
    skipped_test: list[str] = []
    planned_normal: set[str] = set()
    for rec in left_out:
        name = rec["company_name"].strip()
        key = name.casefold()
        emails = [e for e in rec["emails"] if not e.endswith(FIRM_DOMAIN)]
        if TEST_NAME.search(name):
            skipped_test.append(name)
        elif key in existing_keys:
            skipped_existing.append(name)
        elif normal(name) in existing_normal:
            skipped_near.append(name)
        elif normal(name) in planned_normal:
            skipped_near.append(f"{name} (same as another company in this list)")
        elif not emails:
            skipped_no_email.append(name)
        else:
            to_add.append({"company_name": name, "emails": emails})
            planned_normal.add(normal(name))

    print(f"Left-out companies in the list:        {len(left_out)}")
    print(f"Already have a profile (not changed):  {len(skipped_existing)}")
    for name in skipped_existing:
        print(f"    - {name}")
    print(f"Near-match to an existing profile:     {len(skipped_near)} (not added, please review)")
    for name in skipped_near:
        print(f"    - {name}")
    print(f"No client email (not added):           {len(skipped_no_email)}")
    print(f"Test or placeholder names (not added): {len(skipped_test)}")
    for name in skipped_test:
        print(f"    - {name}")
    print(f"To add as profiles with no POC:        {len(to_add)} companies, {len(to_add) * len(templates)} presets")
    print(f"Templates each profile gets:           {', '.join(templates)}")

    if not args.apply:
        print("\nDry run. Nothing was changed. Run with --apply to add these profiles.")
        return

    shutil.copy(PRESETS_PATH, ROOT / "presets.before-left-out.json")
    if PROFILE_EMAILS_PATH.exists():
        shutil.copy(PROFILE_EMAILS_PATH, ROOT / "profile_emails.before-left-out.json")

    for rec in to_add:
        for template in templates:
            presets.append(Preset(rec["company_name"], template, ""))
        profile_emails[rec["company_name"].casefold()] = rec["emails"]
    save_presets(PRESETS_PATH, presets)
    save_profile_emails(PROFILE_EMAILS_PATH, profile_emails)
    print(f"\nAdded {len(to_add)} profiles. Backups: presets.before-left-out.json and profile_emails.before-left-out.json")


if __name__ == "__main__":
    main()
