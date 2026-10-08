"""Export companies and their client email from Zoho CRM into companies.json.

Read-only: this script only makes GET requests to Zoho and never changes any record.
companies.json contains client data, so it is kept out of git by .gitignore.

Rules (see INCLUDE_SINGLE_EMAIL and FIRM_DOMAIN below):
  * A company with one email is included, and that email is its client email.
  * A company with two emails is included only if one of them is @php-cpa.com.sg.
    The other address (not @php-cpa.com.sg) becomes the client email.
  * Companies with more than two emails, or with no usable email, are left out and listed in the summary.

Usage:
  python scripts/export_zoho_companies.py
      Export companies to companies.json (needs the ZOHO_* settings in .env).

  python scripts/export_zoho_companies.py --exchange-code CODE
      One-time setup: turns a Zoho grant code into a refresh token to put in .env.
"""

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = ROOT / ".env"
OUT_PATH = ROOT / "companies.json"

FIRM_DOMAIN = "@php-cpa.com.sg"
INCLUDE_SINGLE_EMAIL = True  # set to False to leave out companies that have only one email
EMAIL_RE = re.compile(r"^[^@\s/]+@[^@\s/]+\.[^@\s/]+$")
CONTACT_FIELDS = "Full_Name,First_Name,Last_Name,Email,Secondary_Email,Account_Name"


def load_env() -> None:
    """Reads .env (KEY=VALUE lines) into os.environ. Values already set in the environment win."""
    if not ENV_PATH.exists():
        return
    for line in ENV_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip())


def require(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        sys.exit(f"Missing {name}. Fill it in .env (see .env.example) and try again.")
    return value


def request_json(url: str, data: dict | None = None, headers: dict | None = None) -> dict:
    body = urllib.parse.urlencode(data).encode() if data is not None else None
    req = urllib.request.Request(url, data=body, headers=headers or {}, method="POST" if data is not None else "GET")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read()
    except urllib.error.HTTPError as exc:
        sys.exit(f"Zoho returned HTTP {exc.code} for {url.split('?')[0]}: {exc.read().decode(errors='replace')}")
    return json.loads(raw) if raw else {}  # Zoho sends an empty body when there are no records


def exchange_code(code: str) -> None:
    accounts = os.environ.get("ZOHO_ACCOUNTS", "https://accounts.zoho.com").rstrip("/")
    result = request_json(f"{accounts}/oauth/v2/token", data={
        "grant_type": "authorization_code",
        "client_id": require("ZOHO_CLIENT_ID"),
        "client_secret": require("ZOHO_CLIENT_SECRET"),
        "code": code,
    })
    refresh = result.get("refresh_token")
    if not refresh:
        sys.exit(f"Zoho did not return a refresh token: {result}")
    print("Refresh token received. Put this in .env as ZOHO_REFRESH_TOKEN (keep it private):")
    print(f"ZOHO_REFRESH_TOKEN={refresh}")


def access_token() -> str:
    accounts = os.environ.get("ZOHO_ACCOUNTS", "https://accounts.zoho.com").rstrip("/")
    result = request_json(f"{accounts}/oauth/v2/token", data={
        "grant_type": "refresh_token",
        "client_id": require("ZOHO_CLIENT_ID"),
        "client_secret": require("ZOHO_CLIENT_SECRET"),
        "refresh_token": require("ZOHO_REFRESH_TOKEN"),
    })
    token = result.get("access_token")
    if not token:
        sys.exit(f"Could not get an access token from Zoho: {result}")
    return token


def fetch_contacts(api: str, token: str) -> list[dict]:
    """Reads every contact. Zoho only allows the first 2000 by page number, so later pages use page_token."""
    rows: list[dict] = []
    page = 1
    page_token = None
    while True:
        params = {"fields": CONTACT_FIELDS, "per_page": 200}
        if page_token:
            params["page_token"] = page_token
        else:
            params["page"] = page
        query = urllib.parse.urlencode(params)
        data = request_json(f"{api}/crm/v8/Contacts?{query}", headers={"Authorization": f"Zoho-oauthtoken {token}"})
        rows.extend(data.get("data", []))
        print(f"  read contacts page {page} ({len(rows)} rows so far)")
        info = data.get("info", {})
        if not info.get("more_records"):
            return rows
        page_token = info.get("next_page_token")
        page += 1
        if not page_token and page > 10:  # without a token, page numbers stop at 2000 records
            sys.exit("Zoho did not return a page token; stopping before the export is complete.")


def build_companies(rows: list[dict]) -> tuple[list[dict], list[str]]:
    """Returns (companies to use, notes explaining what was left out)."""
    names: dict[str, str] = {}
    emails: dict[str, set[str]] = defaultdict(set)
    bad_addresses = 0
    for row in rows:
        account = row.get("Account_Name")
        if not account or not account.get("id"):
            continue
        names[account["id"]] = account.get("name") or ""
        for field in ("Email", "Secondary_Email"):
            address = (row.get(field) or "").strip().lower()
            if not address:
                continue
            if EMAIL_RE.match(address):
                emails[account["id"]].add(address)
            else:
                bad_addresses += 1

    companies: list[dict] = []
    left_out: list[dict] = []  # every company not used above, with its client addresses (no firm addresses)
    notes: list[str] = []
    for account_id, addresses in emails.items():
        name = names[account_id]
        firm = sorted(a for a in addresses if a.endswith(FIRM_DOMAIN))
        other = sorted(a for a in addresses if not a.endswith(FIRM_DOMAIN))
        if len(addresses) == 1 and INCLUDE_SINGLE_EMAIL and not firm:
            companies.append({"company_name": name, "email": other[0], "basis": "one email"})
        elif len(addresses) == 2 and len(firm) == 1:
            companies.append({"company_name": name, "email": other[0],
                              "basis": f"two emails, one {FIRM_DOMAIN} address"})
        else:
            left_out.append({"company_name": name, "emails": other})
            if len(addresses) == 1 and firm:
                notes.append(f"left out (only a {FIRM_DOMAIN} address): {name}")
            elif len(addresses) == 1 and not INCLUDE_SINGLE_EMAIL:
                notes.append(f"left out (one email, single-email rule off): {name}")
            else:
                notes.append(f"left out ({len(addresses)} emails, rule not met): {name}")

    if bad_addresses:
        notes.append(f"{bad_addresses} malformed address(es) ignored (for example 'n/a' or two addresses in one field)")
    companies.sort(key=lambda c: c["company_name"].casefold())
    left_out.sort(key=lambda c: c["company_name"].casefold())
    return companies, notes, left_out


def write_contacts(rows: list[dict]) -> None:
    """Saves every contact's name and email (read-only export) so people can be found by name."""
    people = []
    for row in rows:
        name = (row.get("Full_Name") or " ".join(
            p for p in (row.get("First_Name"), row.get("Last_Name")) if p)).strip()
        account = row.get("Account_Name") or {}
        for field in ("Email", "Secondary_Email"):
            address = (row.get(field) or "").strip().lower()
            if name and address and EMAIL_RE.match(address):
                people.append({"name": name, "email": address, "company_name": account.get("name") or ""})
    people.sort(key=lambda p: (p["name"].casefold(), p["email"]))
    (ROOT / "zoho_contacts.json").write_text(json.dumps(people, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"Saved {len(people)} named contact addresses to zoho_contacts.json")


def export() -> None:
    api = os.environ.get("ZOHO_API", "https://www.zohoapis.com").rstrip("/")
    print("Getting access to Zoho (read-only)...")
    token = access_token()
    print("Reading contacts linked to companies...")
    rows = fetch_contacts(api, token)
    write_contacts(rows)
    companies, notes, left_out = build_companies(rows)
    OUT_PATH.write_text(json.dumps(companies, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    (ROOT / "left_out_companies.json").write_text(json.dumps(left_out, indent=2, ensure_ascii=False) + "\n",
                                                  encoding="utf-8")
    print()
    print(f"Wrote {len(companies)} companies to {OUT_PATH.name}")
    print(f"  one email: {sum(1 for c in companies if c['basis'] == 'one email')}")
    print(f"  two emails with one {FIRM_DOMAIN}: {sum(1 for c in companies if c['basis'] != 'one email')}")
    print(f"Left out: {len(notes)} (listed in left_out.txt)")
    (ROOT / "left_out.txt").write_text("\n".join(notes) + "\n", encoding="utf-8")


def main() -> None:
    load_env()
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--exchange-code", metavar="CODE", help="one-time: turn a Zoho grant code into a refresh token")
    args = parser.parse_args()
    if args.exchange_code:
        exchange_code(args.exchange_code)
    else:
        export()


if __name__ == "__main__":
    main()
