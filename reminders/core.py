"""Templates, presets, images, HTML building, address checks and logs. No UI and no Outlook here."""

import csv
import datetime as dt
import html as html_lib
import json
import os
import re
import shutil
import sys
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path
from string import Formatter

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
if getattr(sys, "frozen", False):
    # Installed copy: the program files are read-only, so data goes in the user's own folder
    ROOT = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "ClientRemindersData"
    ROOT.mkdir(parents=True, exist_ok=True)
else:
    ROOT = Path(__file__).resolve().parent.parent
LOG_PATH = ROOT / "sent_log.csv"
DRAFTS_LOG = ROOT / "drafts_log.csv"
TEMPLATES_PATH = ROOT / "templates.json"
PRESETS_PATH = ROOT / "presets.json"
PROFILE_EMAILS_PATH = ROOT / "profile_emails.json"  # extra emails per company, keyed by casefold company name
PERSON_NAMES_PATH = ROOT / "person_names.json"      # the person behind each address, keyed by lower-case address
SUBSCRIPTIONS_PATH = ROOT / "subscriptions.json"    # mail subscriptions from Zoho invoices, keyed by casefold company
CHANGE_LOG_PATH = ROOT / "change_log.jsonl"         # every data change made in this install, one JSON object per line
IMAGE_DIR = ROOT / "template_images"
# A pasted image is stored in a template body as [[img:<32 hex chars>]], or [[img:<id>|<link>]] if the image is a link
IMAGE_TOKEN = re.compile(r"\[\[img:([0-9a-f]{32})(?:\|([^\]\s|]+))?\]\]")
PLEASE_FILL_IN = re.compile(r"please fill in", re.IGNORECASE)


@dataclass
class Template:
    name: str
    subject: str
    body: str
    to: list[str] = field(default_factory=list)
    cc: list[str] = field(default_factory=list)
    require_attachment: bool = False  # if True, an email using this template needs at least one attachment


@dataclass
class Preset:
    company_name: str
    template_name: str
    email: str = ""  # the client's address(es), separated by ';'. Empty in presets saved before this field existed.


# Data files shipped inside the installed program. They are copied to ROOT the first time the program runs.
SEED_FILES = [
    "templates.json", "presets.json", "profile_emails.json",
    "companies.json", "left_out_companies.json", "drafts_log.csv", "sent_log.csv", "person_names.json",
    "subscriptions.json",
]


def seed_data_folder() -> None:
    """Installed copy only: copies the shipped data into the data folder, never overwriting what is there."""
    if not getattr(sys, "frozen", False):
        return
    shipped = Path(getattr(sys, "_MEIPASS", "")) / "seed"
    for name in SEED_FILES:
        if (shipped / name).exists() and not (ROOT / name).exists():
            shutil.copy(shipped / name, ROOT / name)
    shipped_images = shipped / "template_images"
    if shipped_images.exists():
        IMAGE_DIR.mkdir(exist_ok=True)
        for image in shipped_images.iterdir():
            if not (IMAGE_DIR / image.name).exists():
                shutil.copy(image, IMAGE_DIR / image.name)


def load_templates(path: Path) -> list[Template]:
    data = json.loads(path.read_text(encoding="utf-8"))
    return [Template(**item) for item in data]


def save_templates(path: Path, templates: list[Template]) -> None:
    data = [asdict(t) for t in templates]
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def load_presets(path: Path = PRESETS_PATH) -> list[Preset]:
    if not path.exists():
        return []
    return [Preset(**item) for item in json.loads(path.read_text(encoding="utf-8"))]


def load_profile_emails(path: Path = PROFILE_EMAILS_PATH) -> dict[str, list[str]]:
    return load_json_map(path)


def save_profile_emails(path: Path, data: dict[str, list[str]]) -> None:
    save_json_map(path, data)


def load_json_map(path: Path) -> dict:
    """A small JSON file holding one dictionary, or an empty one if the file doesn't exist yet."""
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def save_json_map(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def save_presets(path: Path, presets: list[Preset]) -> None:
    data = [asdict(p) for p in presets]
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def _merge_profile_emails(local: dict[str, list[str]], shipped: dict[str, list[str]]) -> int:
    """Adds any shipped email not already on the local profile. Never removes or reorders what is there."""
    added = 0
    for key, emails in shipped.items():
        existing = local.setdefault(key, [])
        have = {e.casefold() for e in existing}
        for email in emails:
            if email.casefold() not in have:
                existing.append(email)
                have.add(email.casefold())
                added += 1
    return added


def _sync_templates(local_templates: list[Template],
                     shipped_templates: list[Template]) -> tuple[list[Template], int, list[str]]:
    """Templates always match the shipped (central) version, unlike everything else merge_seed_data touches.

    A template that doesn't exist locally yet is added. A template that exists locally but differs from
    the shipped version (whether it was edited locally, or is just an older central version) is replaced
    with the shipped one. A local template that was never shipped centrally (made only on this PC) is
    left alone. Returns the new template list, how many were added, and the names of any that were
    replaced because they differed locally, so the caller can tell the user which ones changed.
    """
    shipped_by_name = {t.name: t for t in shipped_templates}
    result = []
    seen = set()
    replaced = []
    for local in local_templates:
        shipped = shipped_by_name.get(local.name)
        if shipped is not None and shipped != local:
            result.append(shipped)
            replaced.append(local.name)
        else:
            result.append(local)
        seen.add(local.name)
    added = 0
    for shipped in shipped_templates:
        if shipped.name not in seen:
            result.append(shipped)
            added += 1
    return result, added, replaced


def _merge_dict_add_missing(local: dict, shipped: dict) -> int:
    """Adds a shipped key only if it doesn't exist locally yet. Never overwrites an existing local value."""
    added = 0
    for key, value in shipped.items():
        if key not in local:
            local[key] = value
            added += 1
    return added


def merge_seed_data() -> tuple[int, list[str]]:
    """Installed copy only: adds anything new from this version's shipped data into the user's own data.

    Templates always end up matching the shipped (central) version — see _sync_templates. Everything
    else only ever adds what's missing (a new company, a new preset for an existing company, a new email
    or name) — it never changes or removes anything already saved locally, so a person's own edits, POC
    choices and subscription answers are never touched. Safe to run on every start, since repeating it
    changes nothing more once everything shipped is already matched or present. Returns how many items
    were added in total, and the names of any templates that were replaced with the central version.
    """
    if not getattr(sys, "frozen", False):
        return 0, []
    shipped = Path(getattr(sys, "_MEIPASS", "")) / "seed"
    added = 0
    replaced_templates: list[str] = []

    shipped_templates = shipped / "templates.json"
    if shipped_templates.exists() and TEMPLATES_PATH.exists():
        local_templates = load_templates(TEMPLATES_PATH)
        new_templates, template_added, replaced_templates = _sync_templates(
            local_templates, load_templates(shipped_templates))
        if template_added or replaced_templates:
            save_templates(TEMPLATES_PATH, new_templates)
            added += template_added

    shipped_presets = shipped / "presets.json"
    if shipped_presets.exists() and PRESETS_PATH.exists():
        local_presets = load_presets(PRESETS_PATH)
        have = {(p.company_name.strip().casefold(), p.template_name) for p in local_presets}
        new_presets = [p for p in load_presets(shipped_presets)
                       if (p.company_name.strip().casefold(), p.template_name) not in have]
        if new_presets:
            save_presets(PRESETS_PATH, local_presets + new_presets)
            added += len(new_presets)

    shipped_emails = shipped / "profile_emails.json"
    if shipped_emails.exists():
        local_emails = load_json_map(PROFILE_EMAILS_PATH)
        count = _merge_profile_emails(local_emails, load_json_map(shipped_emails))
        if count:
            save_json_map(PROFILE_EMAILS_PATH, local_emails)
            added += count

    for filename, path in (("person_names.json", PERSON_NAMES_PATH), ("subscriptions.json", SUBSCRIPTIONS_PATH)):
        shipped_path = shipped / filename
        if shipped_path.exists():
            local_map = load_json_map(path)
            count = _merge_dict_add_missing(local_map, load_json_map(shipped_path))
            if count:
                save_json_map(path, local_map)
                added += count

    return added, replaced_templates


def is_duplicate_preset(presets: list[Preset], company: str, template: str, ignore: int | None = None) -> bool:
    """True if another preset has the same company name and template name (ignoring capitals and spaces)."""
    wanted = (company.strip().casefold(), template.strip().casefold())
    return any(
        i != ignore and (p.company_name.strip().casefold(), p.template_name.strip().casefold()) == wanted
        for i, p in enumerate(presets)
    )


def known_companies(presets: list[Preset], log_path: Path = DRAFTS_LOG) -> list[str]:
    """Every company name from the presets and the draft log, sorted, with no repeats."""
    names: dict[str, str] = {}
    for p in presets:
        names.setdefault(p.company_name.strip().casefold(), p.company_name.strip())
    if log_path.exists():
        with log_path.open(newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                name = (row.get("client_name") or "").strip()
                if name:
                    names.setdefault(name.casefold(), name)
    return sorted(names.values(), key=str.casefold)


def last_recipients(company: str, log_path: Path = DRAFTS_LOG) -> str:
    """The To addresses from the most recent draft for this company, or '' if there is none."""
    if not log_path.exists():
        return ""
    found = ""
    with log_path.open(newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):  # the log is in time order, so the last match is the newest
            if (row.get("client_name") or "").strip().casefold() == company.strip().casefold():
                found = (row.get("to") or "").strip()
    return found


def replace_please_fill_in(subject: str, company: str) -> str:
    """Replace every 'please fill in' (any capitals) with the company name."""
    return PLEASE_FILL_IN.sub(lambda _m: company, subject)


def placeholders(*texts: str) -> list[str]:
    """Names of every {placeholder} used in the given texts, in first-seen order."""
    seen: list[str] = []
    for text in texts:
        for _, name, _, _ in Formatter().parse(text):
            if name and name not in seen:
                seen.append(name)
    return seen


def render(text: str, values: dict[str, str]) -> str:
    return text.format_map(values)


def parse_addresses(text: str) -> list[str]:
    """Split on ';' or ',' and drop blanks. Does not validate."""
    return [a.strip() for a in re.split(r"[;,]", text) if a.strip()]


def invalid_addresses(addresses: list[str]) -> list[str]:
    return [a for a in addresses if not EMAIL_RE.match(a)]


def new_image_id() -> str:
    return uuid.uuid4().hex


def image_path(image_id: str) -> Path:
    return IMAGE_DIR / f"{image_id}.png"


def _text_html(text: str) -> str:
    return html_lib.escape(text).replace("\n", "<br>")


def body_to_html(body: str) -> tuple[str, list[str]]:
    """Turn a stored body (text plus [[img:id]] markers) into HTML.

    Returns the HTML and the ids of the images it refers to. Each image is
    referenced as cid:<id>, and the sender attaches the file with that id.
    """
    parts: list[str] = []
    used: list[str] = []
    pos = 0
    for m in IMAGE_TOKEN.finditer(body):
        parts.append(_text_html(body[pos:m.start()]))
        if m.group(1) not in used:
            used.append(m.group(1))
        img = f'<img src="cid:{m.group(1)}" style="max-width:600px">'
        if m.group(2):
            img = f'<a href="{html_lib.escape(m.group(2), quote=True)}">{img}</a>'
        parts.append(img)
        pos = m.end()
    parts.append(_text_html(body[pos:]))
    return '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' + "".join(parts) + "</div>", used


def invalid_images(body: str) -> list[str]:
    """Image ids in the body whose files are missing."""
    return [m.group(1) for m in IMAGE_TOKEN.finditer(body) if not image_path(m.group(1)).exists()]


def valid_link(url: str) -> bool:
    """A link an image can open: http, https or mailto, with no spaces or characters that break the marker."""
    return bool(re.match(r"^(https?://|mailto:)\S+$", url, re.IGNORECASE)) and not any(c in url for c in "]|")


def log_send(template_name: str, to: list[str], cc: list[str], subject: str, path: Path = LOG_PATH) -> None:
    is_new = not path.exists()
    with path.open("a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(["sent_at", "template", "to", "cc", "subject"])
        writer.writerow([dt.datetime.now().isoformat(timespec="seconds"), template_name,
                         "; ".join(to), "; ".join(cc), subject])


def append_change_log(action: str, details: dict, path: Path = CHANGE_LOG_PATH) -> None:
    """Records one data change, for review or export. Never used for presets/profiles/subscriptions reads."""
    entry = {"timestamp": dt.datetime.now().isoformat(timespec="seconds"), "action": action, **details}
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def load_change_log(path: Path = CHANGE_LOG_PATH) -> list[dict]:
    """Every recorded change, oldest first. Lines that are not valid JSON are skipped."""
    if not path.exists():
        return []
    entries = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            entries.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return entries


# Change Log actions that need a template's own content (not in the log) and so are never auto-applied
CHANGE_LOG_MANUAL_ACTIONS = {"template_saved", "template_backfilled", "preset_template_renamed"}


def apply_change_entry(entry: dict, presets: list[Preset], profile_emails: dict, person_names: dict,
                       subscriptions: dict) -> tuple[bool, str]:
    """Replays one Change Log entry against this data. Returns (applied, note on why not if it wasn't).

    Used by both the apply_change_log.py script and the app's own Change Log import screen, so a
    change is replayed the same way everywhere.
    """
    action = entry.get("action", "")
    company = entry.get("company", "")
    key = company.strip().casefold()
    template = entry.get("template", "")
    email = entry.get("email", "")

    if action in ("preset_created", "profile_template_added"):
        if any(p.company_name.strip().casefold() == key and p.template_name == template for p in presets):
            return False, "a preset for this company and template already exists"
        presets.append(Preset(company, template, email))
        return True, ""

    if action == "preset_edited":
        prev_key = entry.get("previous_company", company).strip().casefold()
        prev_template = entry.get("previous_template", template)
        index = next((i for i, p in enumerate(presets)
                      if p.company_name.strip().casefold() == prev_key and p.template_name == prev_template), None)
        if index is None:
            return False, "the preset being edited was not found"
        presets[index] = Preset(company, template, email)
        return True, ""

    if action in ("preset_deleted", "profile_template_removed"):
        before = len(presets)
        presets[:] = [p for p in presets
                      if not (p.company_name.strip().casefold() == key and p.template_name == template)]
        return (len(presets) != before), "" if len(presets) != before else "not found"

    if action == "profile_saved":
        updated = 0
        for p in presets:
            if p.company_name.strip().casefold() == key and p.email != email:
                p.email = email
                updated += 1
        return updated > 0, "" if updated else "no matching presets to update"

    if action == "profile_email_added":
        extras = profile_emails.setdefault(key, [])
        if email.casefold() in {e.casefold() for e in extras}:
            return False, "already on the profile"
        extras.append(email)
        return True, ""

    if action == "profile_email_removed":
        extras = profile_emails.get(key, [])
        before = len(extras)
        extras[:] = [e for e in extras if e.casefold() != email.casefold()]
        return (len(extras) != before), "" if len(extras) != before else "not found"

    if action == "poc_set":
        old = next((p.email.strip() for p in presets
                    if p.company_name.strip().casefold() == key and p.email.strip()), "")
        extras = profile_emails.setdefault(key, [])
        extras[:] = [e for e in extras if e.casefold() != email.casefold()]
        if old and old.casefold() != email.casefold() and old.casefold() not in {e.casefold() for e in extras}:
            extras.append(old)
        changed = False
        for p in presets:
            if p.company_name.strip().casefold() == key:
                p.email = email
                changed = True
        return changed, "" if changed else "no presets for this company"

    if action == "person_name_set":
        person_names[email.strip().lower()] = entry.get("name", "").strip()
        return True, ""

    if action == "person_name_removed":
        return bool(person_names.pop(email.strip().lower(), None)), "no name was set"

    if action == "subscription_status_set":
        status = entry.get("status", "")
        record = subscriptions.setdefault(key, {"company_name": company, "status": status, "invoices": []})
        record["company_name"] = company
        record["status"] = status
        if status == "active" and entry.get("end_date"):
            record["end_date"] = entry["end_date"]
        else:
            record.pop("end_date", None)
        return True, ""

    if action == "subscription_cleared":
        return bool(subscriptions.pop(key, None)), "no subscription was recorded"

    if action == "profile_deleted":
        before = len(presets)
        presets[:] = [p for p in presets if p.company_name.strip().casefold() != key]
        profile_emails.pop(key, None)
        return (len(presets) != before), "" if len(presets) != before else "no presets for this company"

    if action == "template_deleted":
        before = len(presets)
        presets[:] = [p for p in presets if p.template_name != template]
        return (len(presets) != before), "" if len(presets) != before else "no presets used this template"

    return False, "not an auto-applied action type"


def log_draft(client_name: str, template_name: str, to: list[str], cc: list[str], subject: str,
              path: Path = DRAFTS_LOG) -> None:
    """One row per email saved as a draft, with the client (company) name it belongs to."""
    is_new = not path.exists()
    with path.open("a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(["saved_at", "client_name", "template", "to", "cc", "subject"])
        writer.writerow([dt.datetime.now().isoformat(timespec="seconds"), client_name, template_name,
                         "; ".join(to), "; ".join(cc), subject])
