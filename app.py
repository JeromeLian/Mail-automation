"""Client reminder sender: templates, presets, and saving drafts to Outlook classic."""

import calendar
import datetime as dt
import shutil
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from PIL import Image, ImageGrab

from reminders.core import (
    CHANGE_LOG_MANUAL_ACTIONS,
    IMAGE_DIR,
    IMAGE_TOKEN,
    append_change_log,
    apply_change_entry,
    CHANGE_LOG_PATH,
    PERSON_NAMES_PATH,
    PRESETS_PATH,
    PROFILE_EMAILS_PATH,
    ROOT,
    SUBSCRIPTIONS_PATH,
    Preset,
    Template,
    body_to_html,
    image_path,
    invalid_addresses,
    invalid_images,
    is_duplicate_preset,
    known_companies,
    last_recipients,
    load_change_log,
    load_json_map,
    load_presets,
    load_profile_emails,
    load_templates,
    merge_seed_data,
    log_draft,
    new_image_id,
    parse_addresses,
    placeholders,
    render,
    replace_please_fill_in,
    save_json_map,
    save_presets,
    save_profile_emails,
    save_templates,
    seed_data_folder,
    valid_link,
)
from reminders.outlook import DRAFT_ACCOUNT, save_draft
from reminders.version import VERSION

MAX_EMAILS = 50
MAX_IMAGE_WIDTH = 600
IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".bmp")

# Templates that don't need an attachment can be drafted for several clients in one go
MULTI_DRAFT_TEMPLATES = {"Mail Collection Reminder"}

def format_date(iso: str) -> str:
    """An ISO date (YYYY-MM-DD) as dd/mm/yyyy. Returns the input unchanged if it cannot be parsed."""
    try:
        return dt.date.fromisoformat(iso).strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return iso or ""


def parse_date(text: str) -> str | None:
    """dd/mm/yyyy (or d/m/yyyy) as an ISO date. Returns None if the text is not a real date."""
    try:
        return dt.datetime.strptime(text.strip(), "%d/%m/%Y").date().isoformat()
    except ValueError:
        return None


def describe_change(entry: dict) -> str:
    """One plain-language line for a Change Log entry."""
    action = entry.get("action", "")
    company = entry.get("company", "")
    template = entry.get("template", "")
    email = entry.get("email", "")
    name = entry.get("name", "")
    status = entry.get("status", "")
    end_date = entry.get("end_date", "")
    until = f" (ends {format_date(end_date)})" if status == "active" and end_date else ""
    descriptions = {
        "template_deleted": f"Deleted template '{entry.get('name', '')}' "
                            f"({entry.get('presets_removed', 0)} preset(s) removed with it)",
        "template_saved": f"{'Updated' if entry.get('action_type') == 'updated' else 'Created'} "
                          f"template '{entry.get('name', '')}'",
        "template_backfilled": f"Added template '{template}' to {entry.get('companies_added', 0)} "
                               "existing companies",
        "preset_template_renamed": f"Renamed template '{entry.get('old', '')}' to '{entry.get('new', '')}' "
                                   f"({entry.get('count', 0)} preset(s) updated)",
        "profile_saved": f"Saved profile for {company} ({email}): {entry.get('created', 0)} preset(s) created, "
                         f"{entry.get('updated', 0)} updated",
        "preset_created": f"Created preset: {company} / {template} ({email})",
        "preset_edited": f"Edited preset: {entry.get('previous_company', '')} / "
                         f"{entry.get('previous_template', '')} -> {company} / {template} ({email})",
        "preset_deleted": f"Deleted preset: {company} / {template}",
        "profile_template_added": f"Added template '{template}' to {company}",
        "profile_template_removed": f"Removed template '{template}' from {company}",
        "profile_deleted": f"Deleted profile: {company} ({entry.get('preset_count', 0)} preset(s))",
        "profile_email_added": f"Added email to {company}: {email}",
        "profile_email_removed": f"Removed email from {company}: {email}",
        "poc_set": f"Set the POC for {company}: {email}",
        "person_name_set": f"Set the name for {email}: {name}",
        "person_name_removed": f"Removed the name for {email}",
        "subscription_status_set": f"Set the mail subscription for {company}: {status}{until}",
        "subscription_cleared": f"Cleared the mail subscription for {company}",
    }
    return descriptions.get(action, f"{action}: {entry}")


# Mail subscriptions: the template that asks about a client's subscription before drafting
SUBSCRIPTION_TEMPLATE = "Mail Collection Reminder"
SUBSCRIPTION_LABELS = {
    "recorded": "Mail subscription recorded (not yet checked)",
    "active": "Active mail subscription",
    "inactive": "Previous mail subscription, not active",
}

# Text shown in the Guide window. Lines starting with "## " are headings.
GUIDE_TEXT = """## Start here
Client Reminders writes reminder emails and saves them as drafts in Outlook classic. It never sends an email.
Check each draft in Outlook, then press Send yourself.

Outlook classic (the desktop app) must be installed and signed in. The new Outlook does not work.

## The start screen
Templates: Create template, Use template, Edit template, Delete template.
Presets: Use preset, Create preset, Profiles, Edit preset, Delete preset.

## Templates
A template is a reusable email: its subject, body and default CC. Words in braces, such as {client_name}, are blanks.
Each blank gets its own box when you use the template.
- Paste an image with Ctrl+V, or press Add image from file.
- Right-click an image and choose Set link to make it open a web address when clicked.
- Tick "Require at least one attachment" if the email must always include a file.

## Presets and profiles
A preset is one company paired with one template.
A profile is a company with all of its presets. It holds the company's point of contact (POC), the address drafts go to.
Other emails for the company can be kept too, and any of them can be made the POC.
A company with no POC shows "no POC set" and can't be used until a POC is chosen. The app offers to set one straight away.
Create preset: enter a company name and a client email. This creates a preset for every template.

## Use preset
1. Type part of a company name, or a person's name, then click the company. A person's name finds their company.
2. Pick one of the templates saved for that company.
3. Press Continue. The email opens with the POC filled in.
4. Check the email, then press Save as draft.

## Draft multiple emails
Available for templates that don't need an attachment. After you pick a preset, the app asks whether you want to draft for several clients.
- Find existing client: search, then click a client. Tick Select multiple to pick several at once. They go into the Selected clients list. Press Add selected to add them.
- Enter manually: type a company and email, then press Add client.
- Clients to draft for: the list of clients. Remove selected takes clients off. Clear all (keep original) keeps only the first client.
- Start drafting works through the list, one draft at a time.

## Saving drafts
Save as draft puts the email in the Drafts folder of the Outlook account set by your administrator. Nothing is sent.
If the company is new for that template, the app asks whether to keep it as a preset.
Any required attachments must be added first.

## Profiles screen
Lists every company. Search by name, sort by name, and use the POC filter: All companies, With POC, or No POC.
Open a company to set or add its emails, add or remove templates, or delete the profile.

## Lists and filters
The search box, the POC filter and the sort order work the same way in the Profiles list, Use preset, and Find existing client.

## Tips
- Double-click a client in Find existing client to add it, unless Select multiple is ticked.
- Every draft is recorded in drafts_log.csv in your data folder.
- Your data is kept in %LOCALAPPDATA%\\ClientRemindersData. Updating the program does not change it.
- If a client doesn't appear, check the search box and the POC filter.
"""

# Filter for the profile lists: show every company, only those with a POC, or only those without one
POC_FILTERS = ("All companies", "With POC", "No POC")

# Profiles list filter: narrows the list by mail subscription status
SUBSCRIPTION_FILTERS = ("All companies", "Active subscription", "Past subscription, not active", "No subscription")

PICK_TITLES = {
    "edit": ("Edit template", "Edit"),
    "delete": ("Delete template", "Delete"),
    "edit_preset": ("Edit preset", "Edit"),
    "delete_preset": ("Delete preset", "Delete"),
}


class ReminderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Client Reminders")
        self._set_start_geometry(900, 780)
        # Installed copy: brings in anything new from this update (new companies, presets, emails, names).
        # Templates always end up matching the central version; everything else never touches anything
        # already saved locally. Does nothing when running from source.
        self.updates_added, self.templates_replaced = merge_seed_data()
        self.templates_path = ROOT / "templates.json"
        self.templates: dict[str, Template] = {}
        self.presets: list[Preset] = []
        self.fill_entries: dict[str, ttk.Entry] = {}
        self.photos: dict[str, tk.PhotoImage] = {}  # keeps images alive; key is the Tk image name
        self.photo_ids: dict[str, str] = {}         # Tk image name -> stored image id
        self.photo_links: dict[str, str] = {}       # Tk image name -> web address the image opens
        self.company_names: list[str] = []          # every known company, for the Use preset search
        self.company_template_names: list[str] = []  # templates with a preset for the chosen company
        self.attachments: list[str] = []             # file paths attached to the email being drafted
        self.total = 0  # emails planned in this session
        self.done = 0   # emails already saved as drafts in this session
        self.editing_name: str | None = None  # set while the create form is editing an existing template
        self.pick_mode = "edit"

        self.home = ttk.Frame(self, padding=20)
        self.setup_frame = ttk.Frame(self, padding=20)
        self.pick_frame = ttk.Frame(self, padding=20)
        self.use_preset_frame = ttk.Frame(self, padding=20)
        self.profiles_frame = ttk.Frame(self, padding=20)
        self.profile_frame = ttk.Frame(self, padding=20)
        self.batch_frame = ttk.Frame(self, padding=20)
        self.batch_clients: list[tuple[str, str]] = []    # (company, email) clients for multi-draft
        self.pending_names: list[str] = []                # companies picked with Select multiple, not yet added
        self.session_clients: list[tuple[str, str]] = []  # clients for the drafting session in progress
        self.session_template = ""
        self.use_frame = ttk.Frame(self, padding=10)
        self.create_frame = ttk.Frame(self, padding=10)
        self.profile_emails: dict[str, list[str]] = load_profile_emails(PROFILE_EMAILS_PATH)
        self.person_names: dict[str, str] = load_json_map(PERSON_NAMES_PATH)  # lower-case address -> name
        self.subscriptions: dict[str, dict] = load_json_map(SUBSCRIPTIONS_PATH)  # casefold company -> record
        self.profile_email_shown: list[str] = []  # emails listed in the open profile, POC first
        self.profile_keys: list[str] = []         # casefold company name for each line in the profiles list
        self.profile_templates: list[str] = []    # template names shown in the open profile
        self.current_profile = ""                 # casefold company name of the open profile
        self.current_company_name = ""
        self._build_home()
        self._build_setup()
        self._build_profiles()
        self._build_profile()
        self._build_batch()
        self._build_pick()
        self._build_use_preset()
        self._build_use()
        self._build_create()
        self.load_templates()
        self.load_presets()
        self.show(self.home)
        if self.updates_added or self.templates_replaced:
            self.after(300, self.show_update_notice)

    def show_update_notice(self) -> None:
        parts = []
        if self.updates_added:
            parts.append(f"Added {self.updates_added} new item(s) from the central data, such as new "
                         "companies or presets. Nothing you've already set yourself was changed.")
        if self.templates_replaced:
            names = "\n".join(f"  - {name}" for name in self.templates_replaced)
            parts.append("These templates were replaced with the central version, since templates always "
                        f"stay in sync:\n{names}\n\nAny local wording you had for them is gone.")
        messagebox.showinfo("Updated", "\n\n".join(parts), parent=self)

    def _set_start_geometry(self, width: int, height: int) -> None:
        """Opens near the top of the screen so the bottom isn't hidden behind the taskbar."""
        top_margin = 5
        bottom_margin = 60  # rough allowance for the taskbar
        height = min(height, self.winfo_screenheight() - top_margin - bottom_margin)
        x = max(0, (self.winfo_screenwidth() - width) // 2)
        self.geometry(f"{width}x{height}+{x}+{top_margin}")

    def show(self, frame: ttk.Frame) -> None:
        for f in (self.home, self.setup_frame, self.pick_frame, self.use_preset_frame,
                  self.profiles_frame, self.profile_frame, self.batch_frame, self.use_frame, self.create_frame):
            f.pack_forget()
        frame.pack(fill="both", expand=True)

    def load_templates(self) -> None:
        self.templates = {t.name: t for t in load_templates(self.templates_path)}
        names = list(self.templates)
        self.setup_box.config(values=names)
        self.template_box.config(values=names)

    def load_presets(self) -> None:
        self.presets = load_presets(PRESETS_PATH)

    @staticmethod
    def preset_label(preset: Preset) -> str:
        return f"{preset.company_name} - {preset.template_name}"

    def preset_index(self, label: str) -> int | None:
        for i, p in enumerate(self.presets):
            if self.preset_label(p) == label:
                return i
        return None

    # ---- Start screen ----------------------------------------------------

    def _build_home(self) -> None:
        top_row = ttk.Frame(self.home)
        top_row.pack(anchor="e", fill="x")
        ttk.Button(top_row, text="Change Log", command=self.show_change_log).pack(side="right", padx=(6, 0))
        ttk.Button(top_row, text="Guide", command=self.show_guide).pack(side="right")
        ttk.Label(self.home, text="Client Reminders", font=("Segoe UI", 18, "bold")).pack(pady=(10, 10))
        ttk.Label(self.home, foreground="firebrick", justify="center",
                  text="This app only works with classic Outlook (the desktop app).\n"
                       "It does not work with the new Outlook. Open classic Outlook before using it.\n"
                       f"Drafts are saved to {DRAFT_ACCOUNT}. Make sure you are signed in to this email in Outlook classic.",
                  ).pack(pady=(0, 20))

        templates = ttk.LabelFrame(self.home, text="Templates", padding=12)
        templates.pack(fill="x", pady=8)
        for text, command in (
            ("Create template", self.open_create),
            ("Use template", self.open_use),
            ("Edit template", lambda: self.open_pick("edit")),
            ("Delete template", lambda: self.open_pick("delete")),
        ):
            ttk.Button(templates, text=text, width=30, command=command).pack(pady=4, ipady=6)

        presets = ttk.LabelFrame(self.home, text="Presets (company name + template)", padding=12)
        presets.pack(fill="x", pady=(24, 8))
        for text, command in (
            ("Use preset", self.open_use_preset),
            ("Create preset", self.create_preset),
            ("Profiles", self.open_profiles),
            ("Edit preset", lambda: self.open_pick("edit_preset")),
            ("Delete preset", lambda: self.open_pick("delete_preset")),
        ):
            ttk.Button(presets, text=text, width=30, command=command).pack(pady=4, ipady=6)

        version_text = VERSION if getattr(sys, "frozen", False) else f"{VERSION} (dev)"
        ttk.Label(self.home, text=f"Version {version_text}  •  Made by Jerome <3", foreground="gray").pack(
            side="bottom", pady=(16, 0))

    def show_change_log(self) -> None:
        """Opens the Change Log: every data change made in this install, newest first."""
        entries = load_change_log(CHANGE_LOG_PATH)
        win = tk.Toplevel(self)
        win.title("Change Log")
        win.geometry("760x560")
        win.transient(self)

        btns = ttk.Frame(win, padding=10)
        btns.pack(side="bottom", fill="x")
        ttk.Button(btns, text="Export...", command=self.export_change_log).pack(side="left")
        ttk.Button(btns, text="Clear log...", command=lambda: self.clear_change_log(win)).pack(side="left", padx=6)
        ttk.Button(btns, text="Import log...", command=self.import_change_log).pack(side="left", padx=6)
        ttk.Button(btns, text="Close", command=win.destroy).pack(side="right")

        frame = ttk.Frame(win, padding=(10, 10, 10, 0))
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, foreground="gray",
                  text=f"{len(entries)} change(s) recorded on this PC. Newest first. "
                       "Export this to send it to IT.").pack(anchor="w", pady=(0, 6))
        text = tk.Text(frame, wrap="word", font=("Segoe UI", 10))
        scroll = ttk.Scrollbar(frame, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        text.pack(side="left", fill="both", expand=True)
        if not entries:
            text.insert("1.0", "No changes recorded yet.")
        else:
            for entry in reversed(entries):
                text.insert("end", f"{entry.get('timestamp', '')}   {describe_change(entry)}\n")
        text.configure(state="disabled")

    def export_change_log(self) -> None:
        if not CHANGE_LOG_PATH.exists() or not CHANGE_LOG_PATH.read_text(encoding="utf-8").strip():
            messagebox.showinfo("Nothing to export", "No changes have been recorded yet.", parent=self)
            return
        dest = filedialog.asksaveasfilename(
            parent=self, title="Export change log", defaultextension=".jsonl",
            initialfile=f"change_log_{dt.date.today().isoformat()}.jsonl",
            filetypes=[("Change log", "*.jsonl"), ("All files", "*.*")])
        if not dest:
            return
        shutil.copy(CHANGE_LOG_PATH, dest)
        messagebox.showinfo("Exported", f"Saved to:\n{dest}\n\n"
                            "Send this file to IT so they can apply or review the changes.", parent=self)

    def clear_change_log(self, parent: tk.Toplevel) -> None:
        if not CHANGE_LOG_PATH.exists() or not CHANGE_LOG_PATH.read_text(encoding="utf-8").strip():
            messagebox.showinfo("Nothing to clear", "The log is already empty.", parent=parent)
            return
        if not messagebox.askyesno("Clear log",
                                   "Clear the change log on this PC?\n\n"
                                   "Export it first if IT still needs it. This cannot be undone.", parent=parent):
            return
        CHANGE_LOG_PATH.write_text("", encoding="utf-8")
        parent.destroy()
        self.show_change_log()

    def import_change_log(self) -> None:
        """Loads an exported Change Log so each change can be reviewed and approved here."""
        path = filedialog.askopenfilename(parent=self, title="Import change log",
                                          filetypes=[("Change log", "*.jsonl"), ("All files", "*.*")])
        if not path:
            return
        entries = sorted(load_change_log(Path(path)), key=lambda e: e.get("timestamp", ""))
        if not entries:
            messagebox.showinfo("Nothing to import", "That file has no changes in it.", parent=self)
            return
        rows = [{"entry": e, "state": "pending", "note": ""} for e in entries]

        win = tk.Toplevel(self)
        win.title(f"Review changes - {Path(path).name}")
        win.geometry("820x560")
        win.transient(self)

        frame = ttk.Frame(win, padding=(10, 10, 10, 0))
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, foreground="gray", wraplength=780, justify="left",
                  text="Select one or more lines, then Approve selected or Resolve selected. Resolve is for "
                       "template changes, which need to be copied over by hand.").pack(anchor="w", pady=(0, 6))
        listbox = tk.Listbox(frame, selectmode="extended", font=("Segoe UI", 10))
        scroll = ttk.Scrollbar(frame, command=listbox.yview)
        listbox.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True)

        def refresh() -> None:
            listbox.delete(0, "end")
            for row in rows:
                tag = {"pending": "PENDING", "approved": "APPROVED",
                       "resolved": "RESOLVED"}.get(row["state"], f"SKIPPED: {row['note']}")
                listbox.insert("end", f"[{tag}] {row['entry'].get('timestamp', '')}  {describe_change(row['entry'])}")
            pending = sum(1 for r in rows if r["state"] == "pending")
            win.title(f"Review changes - {Path(path).name} ({pending} pending)")

        def apply_rows(indices: list[int]) -> None:
            touched = {"presets": False, "emails": False, "names": False, "subs": False}
            skipped_template = 0
            for i in indices:
                row = rows[i]
                if row["state"] != "pending":
                    continue
                action = row["entry"].get("action", "")
                if action in CHANGE_LOG_MANUAL_ACTIONS:
                    skipped_template += 1
                    continue
                ok, note = apply_change_entry(row["entry"], self.presets, self.profile_emails,
                                              self.person_names, self.subscriptions)
                if ok:
                    row["state"] = "approved"
                    if action in ("preset_created", "profile_template_added", "preset_edited", "preset_deleted",
                                  "profile_template_removed", "profile_saved", "poc_set", "profile_deleted",
                                  "template_deleted"):
                        touched["presets"] = True
                    if action in ("profile_email_added", "profile_email_removed", "poc_set", "profile_deleted"):
                        touched["emails"] = True
                    if action in ("person_name_set", "person_name_removed"):
                        touched["names"] = True
                    if action in ("subscription_status_set", "subscription_cleared"):
                        touched["subs"] = True
                else:
                    row["state"] = "skipped"
                    row["note"] = note
            if touched["presets"]:
                save_presets(PRESETS_PATH, self.presets)
            if touched["emails"]:
                save_profile_emails(PROFILE_EMAILS_PATH, self.profile_emails)
            if touched["names"]:
                save_json_map(PERSON_NAMES_PATH, self.person_names)
            if touched["subs"]:
                save_json_map(SUBSCRIPTIONS_PATH, self.subscriptions)
            refresh()
            if skipped_template:
                messagebox.showinfo("Template changes skipped",
                                    f"{skipped_template} selected line(s) are template changes. "
                                    "Use Resolve selected for those instead.", parent=win)

        def approve_selected() -> None:
            selected = list(listbox.curselection())
            if not selected:
                messagebox.showwarning("Choose changes", "Select one or more lines first.", parent=win)
                return
            apply_rows(selected)

        def approve_all() -> None:
            pending = [i for i, r in enumerate(rows)
                       if r["state"] == "pending" and r["entry"].get("action") not in CHANGE_LOG_MANUAL_ACTIONS]
            if not pending:
                messagebox.showinfo("Nothing to approve", "There are no pending changes to approve.", parent=win)
                return
            if not messagebox.askyesno("Approve all changes?",
                                       f"You are about to approve all {len(pending)} pending change(s) "
                                       "in this log. Continue?", parent=win):
                return
            apply_rows(pending)

        def resolve_selected() -> None:
            selected = list(listbox.curselection())
            if not selected:
                messagebox.showwarning("Choose changes", "Select one or more lines first.", parent=win)
                return
            not_template = 0
            for i in selected:
                row = rows[i]
                if row["state"] != "pending":
                    continue
                if row["entry"].get("action") in CHANGE_LOG_MANUAL_ACTIONS:
                    row["state"] = "resolved"
                else:
                    not_template += 1
            refresh()
            if not_template:
                messagebox.showinfo("Not a template change",
                                    f"{not_template} selected line(s) are not template changes. "
                                    "Use Approve selected for those instead.", parent=win)

        btns = ttk.Frame(win, padding=10)
        btns.pack(side="bottom", fill="x")
        ttk.Button(btns, text="Approve selected", command=approve_selected).pack(side="left")
        ttk.Button(btns, text="Resolve selected", command=resolve_selected).pack(side="left", padx=6)
        ttk.Button(btns, text="Approve all", command=approve_all).pack(side="left", padx=6)
        ttk.Button(btns, text="Close", command=win.destroy).pack(side="right")

        refresh()

    def show_guide(self) -> None:
        """Opens the Guide window: a scrollable, read-only explanation of the app."""
        win = tk.Toplevel(self)
        win.title("Guide - Client Reminders")
        win.geometry("720x620")
        win.transient(self)
        frame = ttk.Frame(win, padding=10)
        frame.pack(fill="both", expand=True)
        text = tk.Text(frame, wrap="word", font=("Segoe UI", 10), padx=12, pady=10)
        scroll = ttk.Scrollbar(frame, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        text.pack(side="left", fill="both", expand=True)
        text.tag_configure("heading", font=("Segoe UI", 12, "bold"), spacing1=12)
        for line in GUIDE_TEXT.splitlines(keepends=True):
            if line.startswith("## "):
                text.insert("end", line[3:], "heading")
            else:
                text.insert("end", line)
        text.configure(state="disabled")
        ttk.Button(frame, text="Close", command=win.destroy).pack(side="bottom", pady=(8, 0))

    def open_create(self) -> None:
        self.editing_name = None
        self.title("Client Reminders")
        self.clear_create_form()
        self.show(self.create_frame)

    def open_use(self) -> None:
        if not self.templates:
            messagebox.showinfo("No templates", "Create a template first.", parent=self)
            return
        if not self.setup_box.get():
            self.setup_box.current(0)
        self.show(self.setup_frame)

    def open_pick(self, mode: str) -> None:
        if mode in ("edit", "delete"):
            if not self.templates:
                messagebox.showinfo("No templates", "There are no templates yet.", parent=self)
                return
            names = list(self.templates)
        else:
            if not self.presets:
                messagebox.showinfo("No presets",
                                    "There are no presets yet. Save one from the email screen "
                                    "by entering a Company Name and saving a draft.", parent=self)
                return
            names = [self.preset_label(p) for p in self.presets]
        self.pick_mode = mode
        title, button_text = PICK_TITLES[mode]
        self.pick_title.config(text=title)
        self.pick_continue_btn.config(text=button_text)
        self.pick_box.config(values=names)
        self.pick_box.set("")
        self.show(self.pick_frame)

    def pick_continue(self) -> None:
        choice = self.pick_box.get()
        if not choice:
            messagebox.showwarning("Choose one", "Pick one from the list first.", parent=self)
            return
        mode = self.pick_mode
        if mode == "edit":
            self.open_edit(choice)
        elif mode == "delete":
            self.delete_template(choice)
        else:
            index = self.preset_index(choice)
            if index is None:
                messagebox.showwarning("Not found", "That preset no longer exists.", parent=self)
                return
            if mode == "edit_preset":
                self.edit_preset(index)
            else:
                self.delete_preset(index)

    def open_edit(self, name: str) -> None:
        t = self.templates[name]
        self.editing_name = name
        self.title(f"Client Reminders - editing {name}")
        self._set_entry(self.new_name, t.name)
        self._set_entry(self.new_to, "; ".join(t.to))
        self._set_entry(self.new_cc, "; ".join(t.cc))
        self._set_entry(self.new_subject, t.subject)
        self.set_body(self.new_body, t.body)
        self.new_require.set(t.require_attachment)
        self.show(self.create_frame)

    def clear_create_form(self) -> None:
        for entry in (self.new_name, self.new_to, self.new_cc, self.new_subject):
            entry.delete(0, "end")
        self.new_body.delete("1.0", "end")
        self.new_require.set(False)

    def delete_template(self, name: str) -> None:
        affected = [p for p in self.presets if p.template_name == name]
        extra = f"\n\n{len(affected)} preset(s) that use it will also be deleted." if affected else ""
        if not messagebox.askyesno("Delete template",
                                   f'Delete the template "{name}"?{extra}\n\nThis cannot be undone.', parent=self):
            return
        remaining = [t for t in self.templates.values() if t.name != name]
        save_templates(self.templates_path, remaining)
        if affected:
            self.presets = [p for p in self.presets if p.template_name != name]
            save_presets(PRESETS_PATH, self.presets)
        self.load_templates()
        self.log_change("template_deleted", name=name, presets_removed=len(affected))
        messagebox.showinfo("Deleted", f'Template "{name}" deleted.', parent=self)
        self.show(self.home)

    # ---- Presets ---------------------------------------------------------

    def create_profile(self, company: str, email: str) -> tuple[int, int]:
        """Makes a preset for every template for this company, using this email.

        Existing presets for the company are kept, and their email is set to this one.
        Returns (created, updated).
        """
        created = updated = 0
        for template in self.templates:
            index = next((i for i, p in enumerate(self.presets)
                          if p.company_name.strip().casefold() == company.casefold()
                          and p.template_name == template), None)
            if index is None:
                self.presets.append(Preset(company, template, email))
                created += 1
            elif self.presets[index].email != email:
                self.presets[index] = Preset(self.presets[index].company_name, template, email)
                updated += 1
        save_presets(PRESETS_PATH, self.presets)
        if created or updated:
            self.log_change("profile_saved", company=company, email=email, created=created, updated=updated)
        return created, updated

    def offer_new_profile(self, company: str, email: str) -> None:
        """After a client is added by hand, offer to make a profile for them if they don't have one."""
        if not self.templates:
            return
        if any(p.company_name.strip().casefold() == company.strip().casefold() for p in self.presets):
            return  # already has a profile
        if messagebox.askyesno(
                "Create profile?",
                f"Create a new profile for {company}?\n\n"
                f"This makes a preset for each of the {len(self.templates)} templates, using {email}. "
                "The company will then appear under Use preset and Profiles.", parent=self):
            created, _updated = self.create_profile(company, email)
            self.company_names = known_companies(self.presets)
            messagebox.showinfo("Profile created", f"Created {created} preset(s) for {company}.", parent=self)

    def create_preset(self) -> None:
        """Company name and email only. Creates a preset for every template for this company."""
        if not self.templates:
            messagebox.showinfo("No templates", "Create a template first, then make a preset for it.", parent=self)
            return
        dlg = tk.Toplevel(self)
        dlg.title("Create preset")
        dlg.transient(self)
        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Company name").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        company_entry = ttk.Entry(frm, width=40)
        company_entry.grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(frm, text="Client email (separate with ;)").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        email_entry = ttk.Entry(frm, width=40)
        email_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(frm, foreground="firebrick", wraplength=460, justify="left", padding=(4, 12, 4, 0),
                  text=f"Notice: this creates a preset for ALL {len(self.templates)} templates for this "
                       "company, so each one can be used from Use preset. If this company already has "
                       "presets, they are kept and their email is changed to the one you enter.",
                  ).grid(row=2, column=0, columnspan=2, sticky="w")

        def save() -> None:
            company = company_entry.get().strip()
            email = "; ".join(parse_addresses(email_entry.get()))
            if not company or not email:
                messagebox.showwarning("Check the preset", "Enter a company name and the client email.", parent=dlg)
                return
            bad = invalid_addresses(parse_addresses(email))
            if bad:
                messagebox.showwarning("Check the preset", "This address looks wrong: " + ", ".join(bad), parent=dlg)
                return
            created, updated = self.create_profile(company, email)
            dlg.destroy()
            messagebox.showinfo("Saved",
                                f"Created {created} preset(s) for {company}.\n"
                                f"Updated the email on {updated} existing preset(s).", parent=self)
            self.show(self.home)

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=4)
        ttk.Button(btns, text="Create presets", command=save).pack(side="left", padx=4)
        dlg.wait_visibility()
        dlg.grab_set()

    def edit_preset(self, index: int) -> None:
        self.preset_form(index)

    def preset_form(self, index: int) -> None:
        """Dialog to edit one preset. Duplicates are refused."""
        preset = self.presets[index]
        dlg = tk.Toplevel(self)
        dlg.title("Edit preset")
        dlg.transient(self)
        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Company name").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        company_entry = ttk.Entry(frm, width=32)
        company_entry.insert(0, preset.company_name)
        company_entry.grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(frm, text="Template").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        template_box = ttk.Combobox(frm, values=list(self.templates), state="readonly", width=30)
        template_box.set(preset.template_name if preset.template_name in self.templates else "")
        template_box.grid(row=1, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(frm, text="Client email (separate with ;)").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        email_entry = ttk.Entry(frm, width=32)
        email_entry.insert(0, preset.email)
        email_entry.grid(row=2, column=1, sticky="ew", padx=4, pady=4)

        def save() -> None:
            company = company_entry.get().strip()
            template = template_box.get()
            email = "; ".join(parse_addresses(email_entry.get()))
            if not company or not template or not email:
                messagebox.showwarning("Check the preset",
                                       "Enter a company name, pick a template and enter the client email.",
                                       parent=dlg)
                return
            bad = invalid_addresses(parse_addresses(email))
            if bad:
                messagebox.showwarning("Check the preset", "This address looks wrong: " + ", ".join(bad), parent=dlg)
                return
            if is_duplicate_preset(self.presets, company, template, ignore=index):
                messagebox.showwarning("Duplicate preset",
                                       "A preset with this company name and template already exists.", parent=dlg)
                return
            before = preset
            self.presets[index] = Preset(company, template, email)
            save_presets(PRESETS_PATH, self.presets)
            self.log_change("preset_edited", company=company, template=template, email=email,
                            previous_company=before.company_name, previous_template=before.template_name,
                            previous_email=before.email)
            dlg.destroy()
            messagebox.showinfo("Saved", "Preset updated.", parent=self)
            self.show(self.home)

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=4)
        ttk.Button(btns, text="Save preset", command=save).pack(side="left", padx=4)
        dlg.wait_visibility()
        dlg.grab_set()

    def delete_preset(self, index: int) -> None:
        label = self.preset_label(self.presets[index])
        if not messagebox.askyesno("Delete preset",
                                   f'Delete the preset "{label}"?\n\n'
                                   "The template and drafts already saved are not affected.", parent=self):
            return
        deleted = self.presets[index]
        del self.presets[index]
        save_presets(PRESETS_PATH, self.presets)
        self.log_change("preset_deleted", company=deleted.company_name, template=deleted.template_name)
        messagebox.showinfo("Deleted", "Preset deleted.", parent=self)
        self.show(self.home)

    def ask_save_preset(self, company: str, template: str) -> str:
        """Returns 'save', 'skip' or 'cancel'."""
        dlg = tk.Toplevel(self)
        dlg.title("Save as preset?")
        dlg.transient(self)
        result = {"choice": "cancel"}

        def pick(choice: str) -> None:
            result["choice"] = choice
            dlg.destroy()

        ttk.Label(dlg, justify="left", padding=16,
                  text=f'Save "{company}" with the template "{template}" as a preset?\n\n'
                       "Presets appear under Use preset on the start screen.").pack()
        row = ttk.Frame(dlg, padding=(16, 0, 16, 16))
        row.pack(fill="x")
        ttk.Button(row, text="Save preset and continue", command=lambda: pick("save")).pack(side="left", padx=4)
        ttk.Button(row, text="Continue without saving", command=lambda: pick("skip")).pack(side="left", padx=4)
        ttk.Button(row, text="Cancel", command=lambda: pick("cancel")).pack(side="right", padx=4)
        dlg.protocol("WM_DELETE_WINDOW", lambda: pick("cancel"))
        dlg.wait_visibility()
        dlg.grab_set()
        self.wait_window(dlg)
        return result["choice"]

    # ---- Draft multiple emails: a list of clients for one template -------

    def _build_batch(self) -> None:
        outer = self.batch_frame
        pad = {"padx": 4, "pady": 4}

        # Bottom buttons stay fixed, outside the scrollable area, packed first so they keep their place
        btns = ttk.Frame(outer)
        btns.pack(side="bottom", fill="x", pady=(12, 0))
        ttk.Button(btns, text="< Back", command=lambda: self.show(self.home)).pack(side="left", padx=4)
        ttk.Button(btns, text="Start drafting", command=self.start_batch_drafting).pack(side="right", padx=4)

        # Everything else scrolls: with Enter manually open, this page can run taller than the window
        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        frm = ttk.Frame(canvas)
        frm_window = canvas.create_window((0, 0), window=frm, anchor="nw")
        frm.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(frm_window, width=e.width))

        def on_mousewheel(event: tk.Event) -> None:
            if isinstance(event.widget, tk.Listbox):
                return  # a listbox under the cursor scrolls itself; don't also scroll the page
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        ttk.Label(frm, text="Draft multiple emails", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        self.batch_title = ttk.Label(frm, foreground="gray")
        self.batch_title.pack(anchor="w", pady=(0, 10))
        self.batch_notice = ttk.Label(frm, foreground="firebrick", wraplength=760, justify="left")
        self.batch_notice.pack(anchor="w", pady=(0, 6))

        ttk.Label(frm, text="Clients to draft for").pack(anchor="w")
        list_row = ttk.Frame(frm)
        list_row.pack(fill="x", **pad)
        self.batch_list = tk.Listbox(list_row, height=6, exportselection=False)
        self.batch_list.pack(side="left", fill="x", expand=True)
        side_btns = ttk.Frame(list_row)
        side_btns.pack(side="left", padx=8, anchor="n")
        ttk.Button(side_btns, text="Remove selected", command=self.remove_batch_client).pack(fill="x")
        self.batch_remove_multi = tk.BooleanVar(value=False)
        ttk.Checkbutton(side_btns, text="Remove multiple", variable=self.batch_remove_multi,
                        command=self.toggle_remove_multi).pack(anchor="w", pady=4)
        ttk.Button(side_btns, text="Clear all (keep original)", command=self.clear_batch_clients).pack(fill="x")

        self.add_client_btn = ttk.Button(frm, text="Add client", command=self.toggle_add_panel)
        self.add_client_btn.pack(anchor="w", pady=(8, 0))

        # Shown when Add client is pressed: two options side by side
        self.add_panel = ttk.Frame(frm)
        self.add_panel.columnconfigure(0, weight=1)
        self.add_panel.columnconfigure(1, weight=1)

        left = ttk.LabelFrame(self.add_panel, text="Find existing client", padding=8)
        left.grid(row=0, column=0, columnspan=2, sticky="nsew")
        self.batch_search = ttk.Entry(left)
        self.batch_search.pack(fill="x")
        self.batch_search.bind("<KeyRelease>", lambda _e: self.refresh_batch_matches())
        batch_filters = ttk.Frame(left)
        batch_filters.pack(fill="x", pady=(4, 0))
        self.batch_poc_filter = ttk.Combobox(batch_filters, values=POC_FILTERS, state="readonly", width=15)
        self.batch_poc_filter.set(POC_FILTERS[0])
        self.batch_poc_filter.pack(side="left")
        self.batch_poc_filter.bind("<<ComboboxSelected>>", lambda _e: self.refresh_batch_matches())
        self.batch_sort_order = ttk.Combobox(batch_filters, values=["Ascending", "Descending"], state="readonly",
                                             width=11)
        self.batch_sort_order.set("Ascending")
        self.batch_sort_order.pack(side="left", padx=6)
        self.batch_sort_order.bind("<<ComboboxSelected>>", lambda _e: self.refresh_batch_matches())
        results = ttk.Frame(left)  # the search results, with the Selected clients list beside them
        results.pack(fill="both", expand=True, pady=(4, 0))
        self.batch_matches = tk.Listbox(results, height=6, exportselection=False)
        self.batch_matches.pack(side="left", fill="both", expand=True)
        self.batch_matches.bind("<Double-Button-1>", self.on_double_click)
        self.batch_matches.bind("<ButtonRelease-1>", self.toggle_pending)
        # Shown only while Select multiple is on: the companies picked so far, beside the search results
        self.pending_frame = ttk.LabelFrame(results, text="Selected clients (0)", padding=6)
        self.pending_list = tk.Listbox(self.pending_frame, height=6, width=28, exportselection=False)
        self.pending_list.pack(fill="x")
        pending_btns = ttk.Frame(self.pending_frame)
        pending_btns.pack(fill="x", pady=(4, 0))
        ttk.Button(pending_btns, text="Remove", command=self.remove_pending).pack(side="left")
        ttk.Button(pending_btns, text="Clear", command=self.clear_pending).pack(side="left", padx=6)
        ttk.Label(left, foreground="gray", wraplength=300, justify="left",
                  text="Click a client to select it, then press Add selected. Double-clicking also adds it.",
                  ).pack(anchor="w")
        add_row = ttk.Frame(left)
        add_row.pack(fill="x", pady=(4, 0))
        self.batch_multi = tk.BooleanVar(value=False)
        ttk.Checkbutton(add_row, text="Select multiple", variable=self.batch_multi,
                        command=self.toggle_multi_select).pack(side="left")
        ttk.Button(add_row, text="Add selected", command=self.add_selected_match).pack(side="right")

        right = ttk.LabelFrame(self.add_panel, text="Enter manually", padding=8)
        right.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Label(right, text="Company name").pack(anchor="w")
        self.manual_company = ttk.Entry(right)
        self.manual_company.pack(fill="x", pady=(0, 6))
        ttk.Label(right, text="Client email (one address)").pack(anchor="w")
        self.manual_email = ttk.Entry(right)
        self.manual_email.pack(fill="x", pady=(0, 6))
        ttk.Label(right, text="Person's name (optional)").pack(anchor="w")
        self.manual_name = ttk.Entry(right)
        self.manual_name.pack(fill="x", pady=(0, 4))
        self.manual_is_poc = tk.StringVar(value="")
        poc_row = ttk.Frame(right)
        poc_row.pack(anchor="w", pady=(0, 6))
        ttk.Label(poc_row, text="Is this person the POC? (needed if a name is entered)").pack(side="left")
        ttk.Radiobutton(poc_row, text="Yes", value="yes", variable=self.manual_is_poc).pack(side="left", padx=4)
        ttk.Radiobutton(poc_row, text="No", value="no", variable=self.manual_is_poc).pack(side="left")
        ttk.Button(right, text="Add client", command=self.add_manual_client).pack(anchor="e", pady=(4, 0))

    def toggle_add_panel(self) -> None:
        if self.add_panel.winfo_ismapped():
            self.add_panel.pack_forget()
        else:
            self.add_panel.pack(fill="x", pady=8, after=self.add_client_btn)

    def email_for_company(self, company: str) -> str:
        key = company.strip().casefold()
        for p in self.presets:
            if p.company_name.strip().casefold() == key and p.email:
                return p.email
        extras = self.profile_emails.get(key, [])
        if extras:  # no POC yet, but an email is already saved on the profile
            return extras[0]
        return last_recipients(company)

    def start_batch(self, template: str, company: str) -> None:
        self.session_template = template
        self.company_names = known_companies(self.presets)
        self.batch_clients = []
        self.pending_names = []
        self.refresh_pending()
        self.batch_title.config(text=f"Template: {template}. Add each client, then press Start drafting.")
        self.add_panel.pack_forget()
        self.add_client_btn.config(text="Add client")
        self.add_batch_client(company, self.email_for_company(company))
        self._set_entry(self.batch_search, "")
        self.refresh_batch_matches()
        self.show(self.batch_frame)

    def add_batch_client(self, company: str, email: str) -> bool:
        if any(c.strip().casefold() == company.strip().casefold() for c, _e in self.batch_clients):
            messagebox.showinfo("Already added", f"{company} is already on the list.", parent=self)
            return False
        self.batch_clients.append((company, email))
        self.refresh_batch_list()
        return True

    def refresh_batch_list(self) -> None:
        self.batch_list.delete(0, "end")
        for company, email in self.batch_clients:
            shown = self.email_label(email) if email else "no email yet"
            self.batch_list.insert("end", f"{company} - {shown}")
        self.refresh_batch_notice()

    def toggle_remove_multi(self) -> None:
        """Remove multiple: lets several clients be highlighted at once, for Remove selected."""
        self.batch_list.config(selectmode="multiple" if self.batch_remove_multi.get() else "browse")

    def clear_batch_clients(self) -> None:
        """Removes every client from the list except the original one (the company the draft started from)."""
        if len(self.batch_clients) <= 1:
            messagebox.showinfo("Nothing to clear", "Only the original client is on the list.", parent=self)
            return
        original = self.batch_clients[0][0]
        if not messagebox.askyesno("Clear all",
                                   f"Remove all clients except {original}?", parent=self):
            return
        self.batch_clients = self.batch_clients[:1]
        self.refresh_batch_list()
        self.refresh_batch_matches()

    def remove_batch_client(self) -> None:
        for index in reversed(self.batch_list.curselection()):
            del self.batch_clients[index]
        self.refresh_batch_list()
        self.refresh_batch_matches()  # a removed company can be found in the search again

    def refresh_batch_matches(self) -> None:
        """Fills the Find existing client list, applying the search box, the POC filter and the sort order."""
        rows = self.company_rows(self.batch_search.get().strip().casefold(), self.batch_poc_filter.get(),
                                 self.batch_sort_order.get() == "Descending")
        on_list = {c.strip().casefold() for c, _e in self.batch_clients}
        rows = [row for row in rows if row[0].strip().casefold() not in on_list]
        self.batch_matches.delete(0, "end")
        self.batch_labels: dict[str, str] = {}  # list line -> company name
        for name, missing, person in rows:
            line = self.company_line(name, missing, person)
            self.batch_labels[line] = name
            self.batch_matches.insert("end", line)

    def toggle_multi_select(self) -> None:
        """Multi-select: clicks pick clients into the Selected clients list. Off: one client at a time."""
        on = self.batch_multi.get()
        self.batch_matches.config(selectmode="multiple" if on else "browse")
        if on:
            self.pending_frame.pack(side="left", fill="y", padx=(8, 0))
        else:
            self.pending_frame.pack_forget()
            self.pending_names = []
            self.refresh_pending()

    def toggle_pending(self, event: tk.Event) -> None:
        """In multi-select, a click picks the client into the list, or takes it out again if already picked."""
        if not self.batch_multi.get():
            return
        index = self.batch_matches.nearest(event.y)
        if index < 0 or index >= self.batch_matches.size():
            return
        self.batch_matches.selection_clear(index)  # the search list stays clear; the picks live in the Selected list
        name = self.batch_labels[self.batch_matches.get(index)]
        if any(n.casefold() == name.casefold() for n in self.pending_names):
            self.pending_names = [n for n in self.pending_names if n.casefold() != name.casefold()]
        else:
            self.pending_names.append(name)
        self.refresh_pending()

    def refresh_pending(self) -> None:
        self.pending_list.delete(0, "end")
        for name in self.pending_names:
            self.pending_list.insert("end", name)
        self.pending_frame.config(text=f"Selected clients ({len(self.pending_names)})")

    def remove_pending(self) -> None:
        for index in reversed(self.pending_list.curselection()):
            del self.pending_names[index]
        self.refresh_pending()

    def clear_pending(self) -> None:
        self.pending_names = []
        self.refresh_pending()

    def on_double_click(self, _event: tk.Event) -> None:
        # In multi-select, a double-click only picks clients; adding is done with Add selected
        if not self.batch_multi.get():
            self.add_selected_match()

    def add_pending_clients(self) -> None:
        """Adds every client in the Selected clients list. Any without a saved email stay in the list."""
        if not self.pending_names:
            messagebox.showwarning("No clients selected",
                                   "Click clients in the list to select them first.", parent=self)
            return
        no_email = []
        added = []
        for company in list(self.pending_names):
            email = self.email_for_company(company)
            if not email:
                no_email.append(company)
            elif self.add_batch_client(company, email):  # already-listed companies are skipped with a message
                added.append(company)
        self.pending_names = no_email
        self.refresh_pending()
        self.refresh_batch_matches()
        if no_email:
            self._set_entry(self.manual_company, no_email[0])
            self.manual_email.focus_set()
            messagebox.showinfo("No email saved",
                                "No email is saved for:\n" + "\n".join(no_email)
                                + "\n\nThese stay in the Selected clients list. Enter each email in the manual box "
                                  "and press Add client, then add them again.", parent=self)
        self.offer_poc_for_added(added)

    def add_selected_match(self, _event: tk.Event | None = None) -> None:
        """Add selected: adds the picked clients in multi-select, or the highlighted client otherwise."""
        if self.batch_multi.get():
            self.add_pending_clients()
        else:
            self.add_highlighted_clients()

    def add_highlighted_clients(self) -> None:
        """Single-select: adds the highlighted client(s). Clients without a saved email go to the manual box."""
        selected = list(self.batch_matches.curselection())
        if not selected:
            messagebox.showwarning("Choose a client", "Click a client to select it first.", parent=self)
            return
        added_rows = []
        added = []
        no_email = []
        for index in selected:
            company = self.batch_labels[self.batch_matches.get(index)]
            email = self.email_for_company(company)
            if not email:
                no_email.append(company)
            elif self.add_batch_client(company, email):
                added_rows.append(index)
                added.append(company)
        for index in sorted(added_rows, reverse=True):  # remove the added rows only; a full rebuild would lag
            self.batch_matches.delete(index)
        if no_email:
            self._set_entry(self.manual_company, no_email[0])
            self.manual_email.focus_set()
            messagebox.showinfo("No email saved",
                                "No email is saved for:\n" + "\n".join(no_email)
                                + "\n\nEnter each email in the manual box and press Add client.", parent=self)
        self.offer_poc_for_added(added)

    def offer_poc_for_added(self, companies: list[str]) -> None:
        """After adding clients from Find existing client, ensures each one has a POC.

        Uses the same "POC not set" prompt as Use preset. For several companies at once, this shows
        one after another rather than a single combined window, so each one gets the exact same prompt,
        the same name requirement, and the same chance to cancel or pick from that company's own emails.
        """
        profiles = self.profile_groups()
        for company in companies:
            if company.strip().casefold() in profiles:
                self.ensure_poc(company)
        self.sync_batch_client_emails(companies)

    def sync_batch_client_emails(self, companies: list[str]) -> None:
        """Updates the Clients to draft for list if a POC was just set (or changed) for any of these companies.

        The email added to the list came from whatever was on hand at the time (the POC, a saved extra
        email, or draft history). If the POC prompt above set or changed the POC, the list should show
        that address instead, since it is now the company's point of contact.
        """
        keys = {c.strip().casefold() for c in companies}
        changed = False
        for i, (company, email) in enumerate(self.batch_clients):
            key = company.strip().casefold()
            if key not in keys:
                continue
            poc = self.poc_for(key)
            if poc and poc != email:
                self.batch_clients[i] = (company, poc)
                changed = True
        if changed:
            self.refresh_batch_list()

    def add_manual_client(self) -> None:
        company = self.manual_company.get().strip()
        email = self.manual_email.get().strip()
        name = self.manual_name.get().strip()
        poc_answer = self.manual_is_poc.get()
        if not company or not email:
            messagebox.showwarning("Check the client", "Enter a company name and the client email.", parent=self)
            return
        if ";" in email or "," in email:
            messagebox.showwarning("One email only", "Enter one email address for this client.", parent=self)
            return
        if invalid_addresses([email]):
            messagebox.showwarning("Check the client", "This address looks wrong.", parent=self)
            return
        if name and not poc_answer:
            messagebox.showwarning("Point of contact",
                                   "Choose Yes or No for whether this person is the POC.", parent=self)
            return
        if not self.add_batch_client(company, email):
            return
        key = company.strip().casefold()
        profile_exists = key in self.profile_groups()
        if name:
            self.save_person_name(email, name)
        if name and profile_exists and poc_answer == "yes":
            self.set_poc(key, email)
        elif name and profile_exists and poc_answer == "no":
            if email.casefold() not in {e.casefold() for e in self.profile_emails_for(key)}:
                self.profile_emails.setdefault(key, []).append(email)
                save_profile_emails(PROFILE_EMAILS_PATH, self.profile_emails)
                self.log_change("profile_email_added", company=company, email=email)
        self._set_entry(self.manual_company, "")
        self._set_entry(self.manual_email, "")
        self._set_entry(self.manual_name, "")
        self.manual_is_poc.set("")
        self.refresh_batch_matches()  # the company may now be hidden from the search
        if not profile_exists:
            self.offer_new_profile(company, email)  # a new profile needs a POC, so this address becomes it

    def start_batch_drafting(self) -> None:
        if not self.batch_clients:
            messagebox.showwarning("No clients", "Add at least one client first.", parent=self)
            return
        without_email = [c for c, e in self.batch_clients if not e]
        if without_email:
            messagebox.showwarning("Missing email",
                                   "These clients have no email yet. Remove them or add their email:\n"
                                   + "\n".join(without_email), parent=self)
            return
        self.session_clients = list(self.batch_clients)
        self.total, self.done = len(self.session_clients), 0
        self.load_client(0)
        self.show(self.use_frame)

    def load_client(self, index: int) -> None:
        """Fill the email screen for the client at this position in the session."""
        company, email = self.session_clients[index]
        if self.session_template == SUBSCRIPTION_TEMPLATE:
            self.check_subscription(company)
            self.refresh_batch_notice()
            if self.subscription_status(company) == "active":
                messagebox.showwarning("Active mail subscription",
                                       f"{company} has an active mail subscription. {SUBSCRIPTION_TEMPLATE} can't "
                                       "be used for them; use Mails Subscription instead. Skipping this client.",
                                       parent=self)
                self.done += 1
                if self.done >= self.total:
                    self.session_clients = []
                    messagebox.showinfo("Finished", f"All {self.total} emails are done.", parent=self)
                    self.show(self.home)
                else:
                    self.load_client(self.done)
                return
        self.clear_attachments()
        self.template_box.set(self.session_template)
        self._set_entry(self.company_entry, company)
        self.on_template_change()
        self._set_entry(self.to_entry, email)
        self.update_progress()

    # ---- Profiles: each company with all of its presets ------------------

    def profile_groups(self) -> dict[str, list[Preset]]:
        """Presets grouped by company. Keys are casefold company names."""
        groups: dict[str, list[Preset]] = {}
        for p in self.presets:
            groups.setdefault(p.company_name.strip().casefold(), []).append(p)
        return groups

    def poc_index(self) -> tuple[dict[str, list[Preset]], dict[str, str]]:
        """Presets grouped by company, and each company's POC, built in one pass for the lists to reuse."""
        groups = self.profile_groups()
        pocs = {key: next((p.email.strip() for p in presets if p.email.strip()), "")
                for key, presets in groups.items()}
        return groups, pocs

    def _build_profiles(self) -> None:
        frm = self.profiles_frame
        frm.columnconfigure(0, weight=1)
        ttk.Label(frm, text="Profiles", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 6))
        ttk.Label(frm, foreground="gray",
                  text="Each company with all of its presets. Double-click a company, or select it and press "
                       "Open. Search matches the company name or a person's name on its profile.",
                  ).pack(anchor="w", pady=(0, 10))
        filters = ttk.Frame(frm)
        filters.pack(fill="x", pady=(0, 8))
        ttk.Label(filters, text="Search").pack(side="left")
        self.profile_search = ttk.Entry(filters, width=40)
        self.profile_search.pack(side="left", padx=6, fill="x", expand=True)
        self.profile_search.bind("<KeyRelease>", lambda _e: self.refresh_profile_list())
        ttk.Label(filters, text="POC").pack(side="left", padx=(12, 0))
        self.poc_filter = ttk.Combobox(filters, values=POC_FILTERS, state="readonly", width=16)
        self.poc_filter.set(POC_FILTERS[0])
        self.poc_filter.pack(side="left", padx=6)
        self.poc_filter.bind("<<ComboboxSelected>>", lambda _e: self.refresh_profile_list())
        ttk.Label(filters, text="Subscription").pack(side="left", padx=(12, 0))
        self.subscription_filter = ttk.Combobox(filters, values=SUBSCRIPTION_FILTERS, state="readonly", width=26)
        self.subscription_filter.set(SUBSCRIPTION_FILTERS[0])
        self.subscription_filter.pack(side="left", padx=6)
        self.subscription_filter.bind("<<ComboboxSelected>>", lambda _e: self.refresh_profile_list())
        ttk.Label(filters, text="Sort by name").pack(side="left", padx=(12, 0))
        self.sort_order = ttk.Combobox(filters, values=["Ascending", "Descending"], state="readonly", width=11)
        self.sort_order.set("Ascending")
        self.sort_order.pack(side="left", padx=4)
        self.sort_order.bind("<<ComboboxSelected>>", lambda _e: self.refresh_profile_list())
        self.profile_count = ttk.Label(frm, foreground="gray")
        self.profile_count.pack(anchor="w", pady=(0, 4))
        self.profile_list = tk.Listbox(frm, height=18, exportselection=False)
        self.profile_list.pack(fill="both", expand=True)
        self.profile_list.bind("<Double-Button-1>", lambda _e: self.open_selected_profile())
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(12, 0))
        ttk.Button(btns, text="< Back", command=lambda: self.show(self.home)).pack(side="left", padx=4)
        ttk.Button(btns, text="Open", command=self.open_selected_profile).pack(side="right", padx=4)

    def open_profiles(self) -> None:
        groups = self.profile_groups()
        if not groups:
            messagebox.showinfo("No profiles", "There are no profiles yet. Use Create preset to make one.", parent=self)
            return
        self.refresh_profile_list()
        self.show(self.profiles_frame)

    def refresh_profile_list(self) -> None:
        """Fills the profiles list, applying the search box, the POC filter and the subscription filter.

        The search matches the company name or the name of a person behind one of its addresses, the
        same rule Use preset and Find existing client use.
        """
        groups, pocs = self.poc_index()
        people = self.people_by_company()
        typed = self.profile_search.get().strip().casefold()
        choice = self.poc_filter.get()
        sub_choice = self.subscription_filter.get()
        # each item: (key, company name, number of templates, has no POC, subscription status or "", matching person)
        items = []
        for key, presets in groups.items():
            name = presets[0].company_name
            missing = not pocs.get(key)
            if choice == "With POC" and missing:
                continue
            if choice == "No POC" and not missing:
                continue
            sub_status = self.subscription_status(key)  # "", "recorded", "active" or "inactive"
            if sub_choice == "Active subscription" and sub_status != "active":
                continue
            if sub_choice == "Past subscription, not active" and sub_status not in ("recorded", "inactive"):
                continue
            if sub_choice == "No subscription" and sub_status != "":
                continue
            person = ""
            if typed and typed not in name.casefold():
                person = next((n for n in people.get(key, []) if typed in n.casefold()), "")
                if not person:
                    continue
            items.append((key, name, len(presets), missing, sub_status, person))
        items.sort(key=lambda i: i[1].casefold(), reverse=self.sort_order.get() == "Descending")

        self.profile_keys = []
        self.profile_list.delete(0, "end")
        for key, name, count, missing, sub_status, person in items:
            line = f"{name}  ({person})" if person else name
            line += f"   ({count} template{'s' if count != 1 else ''})"
            if sub_status:
                line += f"   [{SUBSCRIPTION_LABELS.get(sub_status, sub_status)}]"
            if missing:
                line += "   ⚠ no POC set"
            self.profile_keys.append(key)
            self.profile_list.insert("end", line)
        self.profile_count.config(text=f"Showing {len(self.profile_keys)} of {len(groups)} companies")

    def open_selected_profile(self) -> None:
        selected = self.profile_list.curselection()
        if not selected:
            messagebox.showwarning("Choose a company", "Select a company first.", parent=self)
            return
        self.open_profile(self.profile_keys[selected[0]])

    def _build_profile(self) -> None:
        frm = self.profile_frame
        pad = {"padx": 8, "pady": 6}
        frm.columnconfigure(1, weight=1)

        self.profile_title = ttk.Label(frm, font=("Segoe UI", 16, "bold"))
        self.profile_title.grid(row=0, column=0, sticky="w", pady=(0, 8))
        sub_frame = ttk.Frame(frm)
        sub_frame.grid(row=0, column=1, sticky="e", pady=(0, 8))
        self.profile_subscription = ttk.Label(sub_frame, foreground="gray")
        self.profile_subscription.pack(side="left", padx=(0, 6))
        ttk.Button(sub_frame, text="Edit subscription", command=self.edit_subscription).pack(side="left")

        ttk.Label(frm, text="Point of contact (POC)").grid(row=1, column=0, sticky="w", **pad)
        self.profile_poc = ttk.Label(frm, font=("Segoe UI", 10, "bold"))
        self.profile_poc.grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(frm, text="Emails saved to this client").grid(row=2, column=0, columnspan=2, sticky="w", **pad)
        self.profile_email_list = tk.Listbox(frm, height=5, exportselection=False)
        self.profile_email_list.grid(row=3, column=0, columnspan=2, sticky="nsew", **pad)
        email_btns = ttk.Frame(frm)
        email_btns.grid(row=4, column=0, columnspan=2, sticky="ew", **pad)
        ttk.Button(email_btns, text="Set as POC", command=self.set_selected_poc).pack(side="left", padx=4)
        ttk.Button(email_btns, text="Add email", command=self.add_profile_email).pack(side="left", padx=4)
        ttk.Button(email_btns, text="Remove selected", command=self.remove_profile_email).pack(side="left", padx=4)
        ttk.Button(email_btns, text="Edit name", command=self.edit_profile_name).pack(side="left", padx=4)

        ttk.Label(frm, text="Templates in this profile (one preset each)").grid(
            row=5, column=0, columnspan=2, sticky="w", **pad)
        self.profile_template_list = tk.Listbox(frm, height=6, exportselection=False)
        self.profile_template_list.grid(row=6, column=0, columnspan=2, sticky="nsew", **pad)
        frm.rowconfigure(3, weight=1)
        frm.rowconfigure(6, weight=1)

        side = ttk.Frame(frm)
        side.grid(row=7, column=0, columnspan=2, sticky="ew", **pad)
        ttk.Button(side, text="Add template", command=self.add_profile_template).pack(side="left", padx=4)
        ttk.Button(side, text="Remove selected template", command=self.remove_profile_template).pack(
            side="left", padx=4)

        btns = ttk.Frame(frm)
        btns.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(20, 0))
        ttk.Button(btns, text="< Back to profiles", command=self.open_profiles).pack(side="left", padx=4)
        ttk.Button(btns, text="Delete profile", command=self.delete_profile).pack(side="right", padx=4)

    def open_profile(self, key: str) -> None:
        presets = self.profile_groups().get(key, [])
        if not presets:
            self.open_profiles()
            return
        self.current_profile = key
        self.current_company_name = presets[0].company_name
        self.profile_title.config(text=self.current_company_name)
        status = self.subscription_status(key)
        text = SUBSCRIPTION_LABELS.get(status, "No mail subscription recorded")
        record = self.subscriptions.get(key)
        if status == "active" and record and record.get("end_date"):
            text += f" (until {format_date(record['end_date'])})"
        self.profile_subscription.config(text=text)
        self.refresh_profile_emails()
        self.refresh_profile_templates()
        self.show(self.profile_frame)

    def refresh_profile_templates(self) -> None:
        self.profile_templates = [p.template_name for p in self.presets
                                  if p.company_name.strip().casefold() == self.current_profile]
        self.profile_template_list.delete(0, "end")
        for name in self.profile_templates:
            self.profile_template_list.insert("end", name)

    def subscription_status(self, company: str) -> str:
        """'recorded', 'active', 'inactive', or '' if the company has no recorded mail subscription.

        An active subscription whose end date has passed is reported as 'recorded', so it is asked
        about again, even though the saved record is only updated once Mail Collection Reminder is used.
        """
        record = self.subscriptions.get(company.strip().casefold())
        if not record:
            return ""
        status = record["status"]
        if status == "active" and record.get("end_date"):
            try:
                if dt.date.fromisoformat(record["end_date"]) < dt.date.today():
                    return "recorded"
            except ValueError:
                pass
        return status

    def set_subscription(self, company: str, status: str, end_date: str | None = None) -> None:
        record = self.subscriptions.get(company.strip().casefold())
        if not record:
            return
        record["status"] = status
        if status == "active":
            record["end_date"] = end_date
        else:
            record.pop("end_date", None)
        save_json_map(SUBSCRIPTIONS_PATH, self.subscriptions)
        self.log_change("subscription_status_set", company=self.display_company(company.strip().casefold()),
                        status=status, end_date=end_date or "")

    def ask_subscription_end_date(self, company: str) -> str | None:
        """Opens the calendar to choose the subscription's end date. Returns an ISO date, or None if cancelled."""
        while True:
            end = self.pick_date(self, prompt=f"When does {company}'s mail subscription end?")
            if end is None:
                return None
            if dt.date.fromisoformat(end) < dt.date.today() and not messagebox.askyesno(
                    "Date already passed", f"{format_date(end)} is in the past. Use it anyway?", parent=self):
                continue
            return end

    def pick_date(self, parent: tk.Misc, initial: str | None = None, prompt: str = "Choose a date") -> str | None:
        """A calendar popup to choose a date, so it can't be mistyped. Returns an ISO date, or None if cancelled."""
        today = dt.date.today()
        try:
            start = dt.date.fromisoformat(initial) if initial else today
        except ValueError:
            start = today
        state = {"year": start.year, "month": start.month, "result": None}

        dlg = tk.Toplevel(parent)
        dlg.title("Choose a date")
        dlg.transient(parent)
        frm = ttk.Frame(dlg, padding=12)
        frm.pack()
        ttk.Label(frm, text=prompt, wraplength=260, justify="left").pack(anchor="w", pady=(0, 8))

        header = ttk.Frame(frm)
        header.pack(fill="x")
        ttk.Button(header, text="<", width=3, command=lambda: change_month(-1)).pack(side="left")
        month_label = ttk.Label(header, font=("Segoe UI", 10, "bold"), anchor="center")
        month_label.pack(side="left", expand=True, fill="x")
        ttk.Button(header, text=">", width=3, command=lambda: change_month(1)).pack(side="left")

        grid = ttk.Frame(frm)
        grid.pack(pady=(8, 0))

        def draw() -> None:
            for widget in grid.winfo_children():
                widget.destroy()
            month_label.config(text=f"{calendar.month_name[state['month']]} {state['year']}")
            for col, day_name in enumerate(["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]):
                ttk.Label(grid, text=day_name, width=3, anchor="center").grid(row=0, column=col)
            weeks = calendar.Calendar(firstweekday=0).monthdayscalendar(state["year"], state["month"])
            for row, week in enumerate(weeks, start=1):
                for col, day in enumerate(week):
                    if day == 0:
                        continue
                    ttk.Button(grid, text=str(day), width=3, command=lambda d=day: pick(d)).grid(
                        row=row, column=col, padx=1, pady=1)

        def change_month(delta: int) -> None:
            month = state["month"] + delta
            year = state["year"]
            if month == 0:
                month, year = 12, year - 1
            elif month == 13:
                month, year = 1, year + 1
            state["month"], state["year"] = month, year
            draw()

        def pick(day: int) -> None:
            state["result"] = dt.date(state["year"], state["month"], day).isoformat()
            dlg.destroy()

        draw()
        btns = ttk.Frame(frm)
        btns.pack(pady=(8, 0))
        ttk.Button(btns, text="Go to today",
                   command=lambda: (state.update(year=today.year, month=today.month), draw())).pack(
            side="left", padx=4)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=4)
        dlg.wait_visibility()
        dlg.grab_set()
        self.wait_window(dlg)
        return state["result"]

    def check_subscription(self, company: str) -> None:
        """Mail Collection Reminder: asks about a recorded or active mail subscription before drafting."""
        status = self.subscription_status(company)
        if status == "recorded":
            if messagebox.askyesno("Previous mail subscription",
                                   f"{company} previously had a mail subscription.\n\n"
                                   "Does this client currently have an active subscription?", parent=self):
                end_date = self.ask_subscription_end_date(company)
                if end_date is None:
                    messagebox.showinfo("Not set", "No end date was entered. You'll be asked again next time.",
                                        parent=self)
                    return
                self.set_subscription(company, "active", end_date)
            else:
                self.set_subscription(company, "inactive")
        elif status == "active":
            if not messagebox.askyesno("Active mail subscription",
                                       f"{company} has an active mail subscription.\n\n"
                                       "Please check: is the subscription still active?", parent=self):
                self.set_subscription(company, "inactive")

    def refresh_batch_notice(self) -> None:
        """Lists the selected clients whose previous mail subscription is not active."""
        inactive = [company for company, _email in self.batch_clients
                    if self.subscription_status(company) == "inactive"]
        self.batch_notice.config(text=("Previous mail subscription, not active: " + ", ".join(inactive))
                                 if inactive else "")

    def log_change(self, action: str, **details) -> None:
        """Records a data change for the Change Log. Never blocks the user if it fails to write."""
        try:
            append_change_log(action, details)
        except OSError:
            pass

    def display_company(self, key: str) -> str:
        """The company name as it is actually saved, from a casefold key."""
        return next((p.company_name for p in self.presets if p.company_name.strip().casefold() == key), key)

    def person_name(self, email: str) -> str:
        return self.person_names.get(email.strip().lower(), "")

    def email_label(self, email: str) -> str:
        """The address with the person's name in front, when a name is known."""
        name = self.person_name(email)
        return f"{name} <{email}>" if name else email

    def save_person_name(self, email: str, name: str) -> None:
        self.person_names[email.strip().lower()] = name.strip()
        save_json_map(PERSON_NAMES_PATH, self.person_names)
        self.log_change("person_name_set", email=email.strip().lower(), name=name.strip())

    def ask_name(self, email: str, parent: tk.Misc) -> str | None:
        """Asks for the person's name. Returns None if cancelled."""
        name = simpledialog.askstring("Name", f"Name of the person for {email}:",
                                      initialvalue=self.person_name(email), parent=parent)
        return None if name is None else name.strip()

    def require_name(self, email: str, parent: tk.Misc) -> bool:
        """A POC needs a name. Asks for one until it is given, or until the user cancels."""
        if self.person_name(email):
            return True
        while True:
            name = self.ask_name(email, parent)
            if name is None:
                return False
            if name:
                self.save_person_name(email, name)
                return True
            messagebox.showwarning("Name needed", "A point of contact needs a name.", parent=parent)

    def people_by_company(self) -> dict[str, list[str]]:
        """The names of the people behind each company's addresses, keyed by casefold company name."""
        out: dict[str, list[str]] = {}
        for p in self.presets:
            address = p.email.strip().lower()
            if address and self.person_name(address):
                out.setdefault(p.company_name.strip().casefold(), []).append(self.person_name(address))
        for key, emails in self.profile_emails.items():
            for email in emails:
                if self.person_name(email):
                    out.setdefault(key, []).append(self.person_name(email))
        return out

    def poc_for(self, key: str) -> str:
        """The point of contact for a company: the email on its presets. Empty if none is set."""
        for p in self.presets:
            if p.company_name.strip().casefold() == key and p.email.strip():
                return p.email.strip()
        return ""

    def profile_emails_for(self, key: str) -> list[str]:
        """Every email for a company, the POC first, then the other emails saved to the profile."""
        poc = self.poc_for(key)
        shown = [poc] if poc else []
        for email in self.profile_emails.get(key, []):
            if email.casefold() not in {e.casefold() for e in shown}:
                shown.append(email)
        return shown

    def refresh_profile_emails(self) -> None:
        key = self.current_profile
        poc = self.poc_for(key)
        self.profile_email_shown = self.profile_emails_for(key)
        self.profile_email_list.delete(0, "end")
        for email in self.profile_email_shown:
            is_poc = bool(poc) and email.casefold() == poc.casefold()
            label = self.email_label(email)
            self.profile_email_list.insert("end", f"{label}   (POC)" if is_poc else label)
        self.profile_poc.config(text=poc or "NOT SET", foreground="" if poc else "firebrick")

    def set_poc(self, key: str, email: str) -> None:
        """Makes this email the POC. The previous POC stays on the profile as an extra email."""
        old = self.poc_for(key)
        extras = self.profile_emails.setdefault(key, [])
        extras[:] = [e for e in extras if e.casefold() != email.casefold()]
        if old and old.casefold() != email.casefold() and old.casefold() not in {e.casefold() for e in extras}:
            extras.append(old)
        for p in self.presets:
            if p.company_name.strip().casefold() == key:
                p.email = email
        save_presets(PRESETS_PATH, self.presets)
        save_profile_emails(PROFILE_EMAILS_PATH, self.profile_emails)
        self.log_change("poc_set", company=self.display_company(key), email=email)

    def set_selected_poc(self) -> None:
        selected = self.profile_email_list.curselection()
        if not selected:
            messagebox.showwarning("Choose an email", "Select an email in the list first.", parent=self)
            return
        email = self.profile_email_shown[selected[0]]
        if not self.require_name(email, self):
            return
        self.set_poc(self.current_profile, email)
        self.refresh_profile_emails()
        self.refresh_profile_list_quietly()

    def add_profile_email(self) -> None:
        email = simpledialog.askstring("Add email", "New email address for this client:", parent=self)
        if email is None:
            return
        email = email.strip()
        if not email or invalid_addresses([email]):
            messagebox.showwarning("Check the email", "Enter a valid email address.", parent=self)
            return
        if email.casefold() not in {e.casefold() for e in self.profile_emails_for(self.current_profile)}:
            self.profile_emails.setdefault(self.current_profile, []).append(email)
            save_profile_emails(PROFILE_EMAILS_PATH, self.profile_emails)
            self.log_change("profile_email_added", company=self.current_company_name, email=email)
        name = self.ask_name(email, self)
        if name:
            self.save_person_name(email, name)
        self.refresh_profile_emails()

    def edit_profile_name(self) -> None:
        selected = self.profile_email_list.curselection()
        if not selected:
            messagebox.showwarning("Choose an email", "Select an email in the list first.", parent=self)
            return
        email = self.profile_email_shown[selected[0]]
        name = self.ask_name(email, self)
        if name is None:
            return
        if name:
            self.save_person_name(email, name)
        else:
            self.person_names.pop(email.strip().lower(), None)
            save_json_map(PERSON_NAMES_PATH, self.person_names)
            self.log_change("person_name_removed", email=email.strip().lower())
        self.refresh_profile_emails()

    def remove_profile_email(self) -> None:
        selected = self.profile_email_list.curselection()
        if not selected:
            messagebox.showwarning("Choose an email", "Select an email in the list first.", parent=self)
            return
        email = self.profile_email_shown[selected[0]]
        if email.casefold() == self.poc_for(self.current_profile).casefold():
            messagebox.showwarning("Point of contact",
                                   "This is the POC. Set another email as POC before removing it.", parent=self)
            return
        extras = self.profile_emails.get(self.current_profile, [])
        extras[:] = [e for e in extras if e.casefold() != email.casefold()]
        save_profile_emails(PROFILE_EMAILS_PATH, self.profile_emails)
        self.log_change("profile_email_removed", company=self.current_company_name, email=email)
        self.refresh_profile_emails()

    def refresh_profile_list_quietly(self) -> None:
        """Keeps the profiles list alerts up to date if that list is currently showing."""
        if self.profiles_frame.winfo_ismapped():
            self.open_profiles()

    def ensure_poc(self, company: str) -> bool:
        """Returns True if the company has a POC. Otherwise alerts the user and offers to set one now."""
        key = company.strip().casefold()
        poc = self.poc_for(key)
        if poc and self.person_name(poc):
            return True
        messagebox.showwarning("POC not set",
                               f"{company} needs a point of contact (POC) with a name before it can be used.\n\n"
                               "You need to set the POC to continue.", parent=self)
        if not messagebox.askyesno("Set the profile now?",
                                   "Do you want to set the profile now?", parent=self):
            return False
        return self.choose_poc(key, company)

    def choose_poc(self, key: str, company: str) -> bool:
        """Pick the POC from the company's emails, or type a new one. Returns True if a POC was set."""
        dlg = tk.Toplevel(self)
        dlg.title(f"Set POC for {company}")
        dlg.transient(self)
        result = {"set": False}
        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(0, weight=1)

        ttk.Label(frm, text="Choose an email from this profile, or type a new one.",
                  wraplength=380).grid(row=0, column=0, sticky="w")
        listbox = tk.Listbox(frm, height=6, exportselection=False)
        listbox.grid(row=1, column=0, sticky="ew", pady=(6, 6))
        shown = self.profile_emails_for(key)
        for email in shown:
            listbox.insert("end", self.email_label(email))

        def use_selected() -> None:
            selected = listbox.curselection()
            if not selected:
                messagebox.showwarning("Choose an email", "Select an email in the list first.", parent=dlg)
                return
            email = shown[selected[0]]
            if not self.require_name(email, dlg):
                return
            self.set_poc(key, email)
            result["set"] = True
            dlg.destroy()

        ttk.Label(frm, text="Or enter a new email:").grid(row=2, column=0, sticky="w")
        new_entry = ttk.Entry(frm, width=40)
        new_entry.grid(row=3, column=0, sticky="ew", pady=(4, 6))

        def use_typed() -> None:
            email = new_entry.get().strip()
            if not email or invalid_addresses([email]):
                messagebox.showwarning("Check the email", "Enter a valid email address.", parent=dlg)
                return
            if not self.require_name(email, dlg):
                return
            if email.casefold() not in {e.casefold() for e in self.profile_emails_for(key)}:
                self.profile_emails.setdefault(key, []).append(email)
            self.set_poc(key, email)
            result["set"] = True
            dlg.destroy()

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=4)
        ttk.Button(btns, text="Use typed email", command=use_typed).pack(side="left", padx=4)
        ttk.Button(btns, text="Use selected as POC", command=use_selected).pack(side="left", padx=4)
        dlg.wait_visibility()
        dlg.grab_set()
        self.wait_window(dlg)
        return result["set"]

    def add_profile_template(self) -> None:
        have = set(self.profile_templates)
        options = [t for t in self.templates if t not in have]
        if not options:
            messagebox.showinfo("Nothing to add", "This profile already has every template.", parent=self)
            return
        email = next((p.email for p in self.presets if p.company_name.strip().casefold() == self.current_profile), "")

        dlg = tk.Toplevel(self)
        dlg.title("Add template to profile")
        dlg.transient(self)
        frm = ttk.Frame(dlg, padding=16)
        frm.pack()
        ttk.Label(frm, text="Template").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        box = ttk.Combobox(frm, values=options, state="readonly", width=34)
        box.current(0)
        box.grid(row=0, column=1, padx=4, pady=4)

        def add() -> None:
            self.presets.append(Preset(self.current_company_name, box.get(), email))
            save_presets(PRESETS_PATH, self.presets)
            self.log_change("profile_template_added", company=self.current_company_name, template=box.get())
            dlg.destroy()
            self.refresh_profile_templates()

        btns = ttk.Frame(frm)
        btns.grid(row=1, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=4)
        ttk.Button(btns, text="Add", command=add).pack(side="left", padx=4)
        dlg.wait_visibility()
        dlg.grab_set()

    def remove_profile_template(self) -> None:
        selected = self.profile_template_list.curselection()
        if not selected:
            messagebox.showwarning("Choose a template", "Select a template in the list first.", parent=self)
            return
        name = self.profile_templates[selected[0]]
        if not messagebox.askyesno("Remove template",
                                   f'Remove the "{name}" preset from {self.current_company_name}?', parent=self):
            return
        self.presets = [p for p in self.presets
                        if not (p.company_name.strip().casefold() == self.current_profile and p.template_name == name)]
        save_presets(PRESETS_PATH, self.presets)
        self.log_change("profile_template_removed", company=self.current_company_name, template=name)
        if self.profile_groups().get(self.current_profile):
            self.refresh_profile_templates()
        else:
            self.open_profiles()

    def edit_subscription(self) -> None:
        """Lets the mail subscription status and end date be set or changed by hand on the profile screen."""
        key = self.current_profile
        record = self.subscriptions.get(key, {})
        dlg = tk.Toplevel(self)
        dlg.title(f"Mail subscription for {self.current_company_name}")
        dlg.transient(self)
        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        codes = {"No subscription": "", "Recorded (past, unconfirmed)": "recorded",
                 "Active": "active", "Not active": "inactive"}
        reverse = {v: k for k, v in codes.items()}

        ttk.Label(frm, text="Status").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        status_box = ttk.Combobox(frm, values=list(codes), state="readonly", width=28)
        status_box.set(reverse.get(record.get("status", ""), "No subscription"))
        status_box.grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(frm, text="End date (if active)").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        date_row = ttk.Frame(frm)
        date_row.grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        date_entry = ttk.Entry(date_row, width=12)
        date_entry.insert(0, format_date(record["end_date"]) if record.get("end_date") else "")
        date_entry.pack(side="left")

        def choose_date() -> None:
            picked = self.pick_date(dlg, initial=parse_date(date_entry.get()) or record.get("end_date"),
                                    prompt=f"End date for {self.current_company_name}'s mail subscription")
            if picked:
                self._set_entry(date_entry, format_date(picked))

        ttk.Button(date_row, text="Choose date...", command=choose_date).pack(side="left", padx=(6, 0))
        ttk.Label(frm, foreground="gray", wraplength=320, justify="left",
                  text="dd/mm/yyyy. If the status is Active and this is left blank, you'll be asked for it "
                       "when you save.").grid(row=2, column=0, columnspan=2, sticky="w", padx=4)

        def save() -> None:
            code = codes[status_box.get()]
            if not code:
                self.subscriptions.pop(key, None)
                save_json_map(SUBSCRIPTIONS_PATH, self.subscriptions)
                self.log_change("subscription_cleared", company=self.current_company_name)
                dlg.destroy()
                self.open_profile(key)
                return
            new_record = {"company_name": self.current_company_name, "status": code,
                         "invoices": record.get("invoices", [])}
            if code == "active":
                typed = date_entry.get().strip()
                if typed:
                    end_date = parse_date(typed)
                    if end_date is None:
                        messagebox.showwarning("Check the date",
                                               "Enter the date as dd/mm/yyyy, for example 31/12/2026, "
                                               "or use Choose date.", parent=dlg)
                        return
                else:
                    end_date = self.ask_subscription_end_date(self.current_company_name)
                    if end_date is None:
                        messagebox.showwarning("End date needed",
                                               "An active subscription needs an end date.", parent=dlg)
                        return
                new_record["end_date"] = end_date
            self.subscriptions[key] = new_record
            save_json_map(SUBSCRIPTIONS_PATH, self.subscriptions)
            self.log_change("subscription_status_set", company=self.current_company_name, status=code,
                            end_date=new_record.get("end_date", ""))
            dlg.destroy()
            self.open_profile(key)

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=4)
        ttk.Button(btns, text="Save", command=save).pack(side="left", padx=4)
        dlg.wait_visibility()
        dlg.grab_set()

    def delete_profile(self) -> None:
        count = len(self.profile_templates)
        if not messagebox.askyesno(
                "Delete profile",
                f"Delete {self.current_company_name} and its {count} preset(s)?\n\n"
                "The templates and drafts already saved are not affected. This cannot be undone.", parent=self):
            return
        self.presets = [p for p in self.presets if p.company_name.strip().casefold() != self.current_profile]
        save_presets(PRESETS_PATH, self.presets)
        self.profile_emails.pop(self.current_profile, None)
        save_profile_emails(PROFILE_EMAILS_PATH, self.profile_emails)
        self.log_change("profile_deleted", company=self.current_company_name, preset_count=count)
        messagebox.showinfo("Deleted", "Profile deleted.", parent=self)
        self.open_profiles()

    # ---- Use preset: pick a company, then one of its templates -----------

    def _build_use_preset(self) -> None:
        frm = self.use_preset_frame
        pad = {"padx": 8, "pady": 6}
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Use preset", font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        # Step 1: company search, with the matches listed underneath
        ttk.Label(frm, text="1. Company or person name (type to search)").grid(row=1, column=0, sticky="nw", **pad)
        company_col = ttk.Frame(frm)
        company_col.grid(row=1, column=1, sticky="nsew", **pad)
        company_col.columnconfigure(0, weight=1)
        self.up_company = ttk.Entry(company_col)
        self.up_company.grid(row=0, column=0, sticky="ew")
        self.up_company.bind("<KeyRelease>", lambda _e: self.on_company_typed())
        filters = ttk.Frame(company_col)
        filters.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.up_poc_filter = ttk.Combobox(filters, values=POC_FILTERS, state="readonly", width=16)
        self.up_poc_filter.set(POC_FILTERS[0])
        self.up_poc_filter.pack(side="left")
        self.up_poc_filter.bind("<<ComboboxSelected>>", lambda _e: self.refresh_company_list())
        ttk.Label(filters, text="Sort by name").pack(side="left", padx=(12, 0))
        self.up_sort_order = ttk.Combobox(filters, values=["Ascending", "Descending"], state="readonly", width=11)
        self.up_sort_order.set("Ascending")
        self.up_sort_order.pack(side="left", padx=4)
        self.up_sort_order.bind("<<ComboboxSelected>>", lambda _e: self.refresh_company_list())
        self.company_list = tk.Listbox(company_col, height=8, exportselection=False)
        self.company_list.grid(row=2, column=0, sticky="ew", pady=(4, 0))
        self.company_list.bind("<<ListboxSelect>>", self.pick_company)
        ttk.Label(company_col, foreground="gray", wraplength=330, justify="left",
                  text="Click a match to select it. If you searched for a person, their company is shown with "
                       "the person's name in brackets.").grid(row=3, column=0, sticky="w")

        # Step 2: template search, limited to templates saved for the chosen company
        ttk.Label(frm, text="2. Template (type to search)").grid(row=2, column=0, sticky="nw", **pad)
        template_col = ttk.Frame(frm)
        template_col.grid(row=2, column=1, sticky="nsew", **pad)
        template_col.columnconfigure(0, weight=1)
        self.up_template = ttk.Entry(template_col)
        self.up_template.grid(row=0, column=0, sticky="ew")
        self.up_template.bind("<KeyRelease>", lambda _e: self.refresh_template_list())
        self.template_list = tk.Listbox(template_col, height=5, exportselection=False)
        self.template_list.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.template_list.bind("<<ListboxSelect>>", self.pick_template)
        ttk.Label(template_col, foreground="gray", text="Click a match to select it").grid(
            row=2, column=0, sticky="w")

        self.up_hint = ttk.Label(frm, foreground="gray", wraplength=560, justify="left")
        self.up_hint.grid(row=3, column=0, columnspan=2, sticky="w", **pad)
        frm.rowconfigure(1, weight=1)
        frm.rowconfigure(2, weight=1)

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(20, 0))
        ttk.Button(btns, text="< Back", command=lambda: self.show(self.home)).pack(side="left", padx=4)
        ttk.Button(btns, text="Continue", command=self.continue_use_preset).pack(side="right", padx=4)

    def open_use_preset(self) -> None:
        if not self.templates:
            messagebox.showinfo("No templates", "Create a template first.", parent=self)
            return
        self.company_names = known_companies(self.presets)
        self._set_entry(self.up_company, "")
        self.refresh_company_list()
        self.update_template_choices()
        self.show(self.use_preset_frame)

    def on_company_typed(self) -> None:
        self.refresh_company_list()
        self.update_template_choices()

    def company_rows(self, typed: str, choice: str, descending: bool) -> list[tuple[str, bool, str]]:
        """Companies matching the search, the POC filter and the sort order.

        The search matches the company name or the name of a person behind one of its addresses.
        Each row is (company name, has no POC, the matching person's name or "").
        """
        profiles, pocs = self.poc_index()
        people = self.people_by_company()
        rows = []
        for name in self.company_names:
            key = name.casefold()
            is_profile = key in profiles  # only profiles can be missing a POC
            missing = is_profile and not pocs.get(key)
            if choice == "With POC" and not (is_profile and not missing):
                continue
            if choice == "No POC" and not missing:
                continue
            person = ""
            if typed not in name.casefold():
                person = next((n for n in people.get(key, []) if typed in n.casefold()), None)
                if person is None:
                    continue
            rows.append((name, missing, person or ""))
        rows.sort(key=lambda r: r[0].casefold(), reverse=descending)
        return rows

    @staticmethod
    def company_line(name: str, missing: bool, person: str = "") -> str:
        text = f"{name}  ({person})" if person else name
        return f"{text}   ⚠ no POC set" if missing else text

    def refresh_company_list(self) -> None:
        """Fills the company list, applying the search box, the POC filter and the sort order."""
        rows = self.company_rows(self.up_company.get().strip().casefold(), self.up_poc_filter.get(),
                                 self.up_sort_order.get() == "Descending")
        self.company_list.delete(0, "end")
        self.company_labels: dict[str, str] = {}  # list line -> company name
        for name, missing, person in rows:
            line = self.company_line(name, missing, person)
            self.company_labels[line] = name
            self.company_list.insert("end", line)

    def pick_company(self, _event: tk.Event) -> None:
        selected = self.company_list.curselection()
        if not selected:
            return
        self._set_entry(self.up_company, self.company_labels[self.company_list.get(selected[0])])
        self.update_template_choices()

    def update_template_choices(self) -> None:
        """Show only the templates that have a preset for the company typed in step 1."""
        company = self.up_company.get().strip().casefold()
        names: list[str] = []
        for p in self.presets:
            if p.company_name.strip().casefold() == company and p.template_name in self.templates:
                if p.template_name not in names:
                    names.append(p.template_name)
        self.company_template_names = names
        self._set_entry(self.up_template, names[0] if len(names) == 1 else "")
        self.refresh_template_list()

        if not company:
            hint = "Select a company first. Its saved templates will appear here."
        elif not names:
            hint = ("No templates are saved for this company. Save a draft with this company name "
                    "and choose 'Save preset', or use Use template instead.")
        else:
            hint = "Pick one of the templates saved for this company."
        self.up_hint.config(text=hint)

    def refresh_template_list(self) -> None:
        typed = self.up_template.get().strip().casefold()
        self.template_list.delete(0, "end")
        for name in self.company_template_names:
            if typed in name.casefold():
                self.template_list.insert("end", name)

    def pick_template(self, _event: tk.Event) -> None:
        selected = self.template_list.curselection()
        if selected:
            self._set_entry(self.up_template, self.template_list.get(selected[0]))

    def continue_use_preset(self) -> None:
        company = self.up_company.get().strip()
        template = self.up_template.get().strip()
        if not company:
            messagebox.showwarning("Company name", "Type or pick a company name first.", parent=self)
            return
        if template not in self.company_template_names:
            messagebox.showwarning("Choose a template",
                                   "Pick one of the templates saved for this company.", parent=self)
            return
        if not self.ensure_poc(company):
            return
        if template == SUBSCRIPTION_TEMPLATE:
            self.check_subscription(company)
            if self.subscription_status(company) == "active":
                messagebox.showwarning("Active mail subscription",
                                       f"{company} has an active mail subscription. {SUBSCRIPTION_TEMPLATE} can't "
                                       "be used for them; use Mails Subscription instead, which needs an "
                                       "attachment.", parent=self)
                return
        if template in MULTI_DRAFT_TEMPLATES and messagebox.askyesno(
                "Draft multiple emails?",
                f'"{template}" does not need an attachment.\n\n'
                "Do you want to draft emails for more than one client?", parent=self):
            self.start_batch(template, company)
            return
        self.session_clients = []
        self.total, self.done = 1, 0
        self.template_box.set(template)
        self._set_entry(self.company_entry, company)
        self.clear_attachments()
        self.on_template_change()
        # The preset's email; for presets saved without one, the address used last time for this company
        preset = next((p for p in self.presets if p.company_name.strip().casefold() == company.casefold()
                       and p.template_name == template), None)
        self._set_entry(self.to_entry, preset.email if preset and preset.email else last_recipients(company))
        self.update_progress()
        self.show(self.use_frame)

    # ---- Use template: setup (pick template and count) -------------------

    def _build_setup(self) -> None:
        frm = self.setup_frame
        pad = {"padx": 8, "pady": 8}
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Use template", font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 16))

        ttk.Label(frm, text="Which template?").grid(row=1, column=0, sticky="w", **pad)
        self.setup_box = ttk.Combobox(frm, state="readonly")
        self.setup_box.grid(row=1, column=1, sticky="ew", **pad)

        ttk.Label(frm, text=f"How many emails to draft? (1-{MAX_EMAILS})").grid(
            row=2, column=0, sticky="w", **pad)
        self.count_var = tk.StringVar(value="1")
        ttk.Spinbox(frm, from_=1, to=MAX_EMAILS, textvariable=self.count_var, width=8).grid(
            row=2, column=1, sticky="w", **pad)

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(24, 0))
        ttk.Button(btns, text="< Back", command=lambda: self.show(self.home)).pack(side="left", padx=4)
        ttk.Button(btns, text="Continue", command=self.start_session).pack(side="right", padx=4)

    def start_session(self) -> None:
        name = self.setup_box.get()
        if not name:
            messagebox.showwarning("Choose a template", "Pick a template first.", parent=self)
            return
        try:
            count = int(self.count_var.get())
        except ValueError:
            count = 0
        if not 1 <= count <= MAX_EMAILS:
            messagebox.showwarning("How many emails?",
                                   f"Enter a whole number from 1 to {MAX_EMAILS}.", parent=self)
            return

        self.total, self.done = count, 0
        self.session_clients = []
        self.template_box.set(name)
        self._set_entry(self.company_entry, "")
        self.clear_attachments()
        self.on_template_change()
        self.update_progress()
        self.show(self.use_frame)

    # ---- Edit / delete: pick a template or preset -------------------------

    def _build_pick(self) -> None:
        frm = self.pick_frame
        pad = {"padx": 8, "pady": 8}
        frm.columnconfigure(1, weight=1)

        self.pick_title = ttk.Label(frm, font=("Segoe UI", 16, "bold"))
        self.pick_title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 16))

        ttk.Label(frm, text="Choose").grid(row=1, column=0, sticky="w", **pad)
        self.pick_box = ttk.Combobox(frm, state="readonly")
        self.pick_box.grid(row=1, column=1, sticky="ew", **pad)

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(24, 0))
        ttk.Button(btns, text="< Back", command=lambda: self.show(self.home)).pack(side="left", padx=4)
        self.pick_continue_btn = ttk.Button(btns, command=self.pick_continue)
        self.pick_continue_btn.pack(side="right", padx=4)

    # ---- Use template: drafting one email at a time ----------------------

    def _build_use(self) -> None:
        frm = self.use_frame
        pad = {"padx": 8, "pady": 4}
        frm.columnconfigure(1, weight=1)

        self.progress = ttk.Label(frm, font=("Segoe UI", 12, "bold"))
        self.progress.grid(row=0, column=0, columnspan=2, sticky="w", **pad)

        ttk.Label(frm, text="Template").grid(row=1, column=0, sticky="w", **pad)
        self.template_box = ttk.Combobox(frm, state="readonly")
        self.template_box.grid(row=1, column=1, sticky="ew", **pad)
        self.template_box.bind("<<ComboboxSelected>>", lambda _e: self.on_template_change())

        ttk.Label(frm, text="To: client email (separate with ;)").grid(row=2, column=0, sticky="w", **pad)
        self.to_entry = ttk.Entry(frm)
        self.to_entry.grid(row=2, column=1, sticky="ew", **pad)

        ttk.Label(frm, text="CC (separate with ;)").grid(row=3, column=0, sticky="w", **pad)
        self.cc_entry = ttk.Entry(frm)
        self.cc_entry.grid(row=3, column=1, sticky="ew", **pad)

        self.fields_frame = ttk.LabelFrame(frm, text="Fill in", padding=8)
        self.fields_frame.grid(row=4, column=0, columnspan=2, sticky="ew", **pad)
        self.fields_frame.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Subject").grid(row=5, column=0, sticky="w", **pad)
        self.subject_entry = ttk.Entry(frm)
        self.subject_entry.grid(row=5, column=1, sticky="ew", **pad)

        ttk.Label(frm, text="Body").grid(row=6, column=0, sticky="nw", **pad)
        self.body_text = tk.Text(frm, height=14, wrap="word")
        self.body_text.grid(row=6, column=1, sticky="nsew", **pad)
        self.body_text.bind("<Control-v>", self.paste_into)
        self.body_text.bind("<Button-3>", self.image_menu)
        frm.rowconfigure(6, weight=1)

        # Optional side panel. Its text replaces "please fill in" in the subject.
        side = ttk.LabelFrame(frm, text="Optional", padding=10)
        side.grid(row=1, column=2, rowspan=3, sticky="n", padx=(8, 4), pady=4)
        ttk.Label(side, text="Company Name").pack(anchor="w")
        self.company_entry = ttk.Entry(side, width=26)
        self.company_entry.pack(fill="x", pady=(2, 6))
        ttk.Label(side, foreground="gray", wraplength=190, justify="left",
                  text="If filled in, 'please fill in' in the subject is replaced with this name. "
                       "Saving a draft then offers to keep it as a preset.").pack(anchor="w")

        attach = ttk.LabelFrame(frm, text="Attachments", padding=6)
        attach.grid(row=7, column=0, columnspan=3, sticky="ew", **pad)
        attach.columnconfigure(0, weight=1)
        self.attach_list = tk.Listbox(attach, height=3, exportselection=False)
        self.attach_list.grid(row=0, column=0, sticky="ew")
        attach_btns = ttk.Frame(attach)
        attach_btns.grid(row=0, column=1, sticky="n", padx=(6, 0))
        ttk.Button(attach_btns, text="Add attachment", command=self.add_attachment).pack(fill="x", pady=2)
        ttk.Button(attach_btns, text="Remove selected", command=self.remove_attachment).pack(fill="x", pady=2)

        btns = ttk.Frame(frm)
        btns.grid(row=8, column=0, columnspan=3, sticky="ew", **pad)
        ttk.Button(btns, text="< Stop", command=self.stop_session).pack(side="left", padx=4)

        # Actions for the email (kept apart from the stop button)
        actions = ttk.LabelFrame(btns, text="Email", padding=(6, 2))
        actions.pack(side="right", padx=4)
        ttk.Button(actions, text="Save as draft", command=self.save_as_draft).pack(side="right", padx=4)
        ttk.Button(actions, text="Add image from file",
                   command=lambda: self.add_image_from_file(self.body_text)).pack(side="right", padx=4)
        ttk.Button(actions, text="Fill template", command=self.fill_template).pack(side="right", padx=4)

        self.status = ttk.Label(frm, text="", foreground="gray")
        self.status.grid(row=9, column=0, columnspan=3, sticky="w", **pad)

    def add_attachment(self) -> None:
        paths = filedialog.askopenfilenames(parent=self, title="Choose attachments")
        for path in paths:
            if path not in self.attachments:
                self.attachments.append(path)
        self.refresh_attachment_list()

    def remove_attachment(self) -> None:
        for index in reversed(self.attach_list.curselection()):
            del self.attachments[index]
        self.refresh_attachment_list()

    def clear_attachments(self) -> None:
        self.attachments.clear()
        self.refresh_attachment_list()

    def refresh_attachment_list(self) -> None:
        self.attach_list.delete(0, "end")
        for path in self.attachments:
            self.attach_list.insert("end", Path(path).name)

    def update_progress(self) -> None:
        client = f"{self.session_clients[self.done][0]}, " if self.session_clients else ""
        self.progress.config(text=f"Email {self.done + 1} of {self.total}  ({client}{self.template_box.get()})")

    def stop_session(self) -> None:
        if self.done < self.total and not messagebox.askyesno(
                "Stop?", f"You have saved {self.done} of {self.total} emails. Stop now?", parent=self):
            return
        self.session_clients = []
        self.show(self.home)

    def email_done(self, message: str) -> None:
        self.done += 1
        if self.done >= self.total:
            self.session_clients = []
            messagebox.showinfo("Finished", f"All {self.total} emails are done.", parent=self)
            self.show(self.home)
            return
        self.clear_attachments()
        if self.session_clients:
            self.load_client(self.done)  # the next client on the list
        else:
            self._set_entry(self.company_entry, "")  # the company name is for one email at a time
            self.on_template_change()  # clear the boxes for the next email
            self.update_progress()
        self.status.config(text=f"{message}. Fill in the next email.", foreground="green")

    # ---- Images and rich body text ---------------------------------------

    def paste_into(self, event: tk.Event) -> str | None:
        """Ctrl+V: paste a clipboard image (or copied image files) into the body. Plain text pastes as normal."""
        try:
            clip = ImageGrab.grabclipboard()
        except Exception:
            return None
        if isinstance(clip, Image.Image):
            self.insert_pil(event.widget, clip, "insert")
            return "break"
        if isinstance(clip, list):  # image files copied in File Explorer
            files = [f for f in clip if str(f).lower().endswith(IMAGE_EXTENSIONS)]
            if files:
                for f in files:
                    self.insert_pil(event.widget, Image.open(f), "insert")
                return "break"
        return None

    def add_image_from_file(self, text: tk.Text) -> None:
        path = filedialog.askopenfilename(
            parent=self, title="Choose an image",
            filetypes=[("Images", " ".join("*" + e for e in IMAGE_EXTENSIONS))])
        if not path:
            return
        try:
            self.insert_pil(text, Image.open(path), "insert")
        except Exception as exc:
            messagebox.showerror("Image not added", str(exc), parent=self)

    def insert_pil(self, text: tk.Text, img: Image.Image, index: str) -> None:
        """Save a PIL image as a template image file and show it in the text box."""
        IMAGE_DIR.mkdir(exist_ok=True)
        image_id = new_image_id()
        img = img.convert("RGBA")
        if img.width > MAX_IMAGE_WIDTH:
            img = img.resize((MAX_IMAGE_WIDTH, round(img.height * MAX_IMAGE_WIDTH / img.width)))
        img.save(image_path(image_id), "PNG")
        self.insert_image(text, image_id, index)

    def insert_image(self, text: tk.Text, image_id: str, index: str, link: str = "") -> None:
        path = image_path(image_id)
        if not path.exists():
            text.insert(index, "[image missing]")
            return
        photo = tk.PhotoImage(file=str(path))
        self.photos[str(photo)] = photo
        self.photo_ids[str(photo)] = image_id
        if link:
            self.photo_links[str(photo)] = link
        text.image_create(index, image=photo)

    def set_body(self, text: tk.Text, body: str) -> None:
        """Show a stored body (text plus [[img:id]] markers) with the images in place."""
        text.delete("1.0", "end")
        pos = 0
        for m in IMAGE_TOKEN.finditer(body):
            text.insert("end", body[pos:m.start()])
            self.insert_image(text, m.group(1), "end", m.group(2) or "")
            pos = m.end()
        text.insert("end", body[pos:])

    def image_menu(self, event: tk.Event) -> None:
        """Right-click on an image: set or remove the web address it opens."""
        text = event.widget
        try:
            name = text.image_cget(text.index(f"@{event.x},{event.y}"), "image")
        except tk.TclError:
            return
        if name not in self.photo_ids:
            return
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Set link...", command=lambda: self.set_image_link(name))
        menu.add_command(label="Remove link", command=lambda: self.photo_links.pop(name, None))
        menu.tk_popup(event.x_root, event.y_root)

    def set_image_link(self, name: str) -> None:
        url = simpledialog.askstring(
            "Link for image",
            "Web address the image should open (starts with https://, or mailto:).\n"
            "Leave blank to remove the link.",
            initialvalue=self.photo_links.get(name, ""), parent=self)
        if url is None:
            return
        url = url.strip()
        if not url:
            self.photo_links.pop(name, None)
            return
        if not valid_link(url):
            messagebox.showwarning("Link not set",
                                   "The address must start with https:// (or mailto:) and have no spaces.",
                                   parent=self)
            return
        self.photo_links[name] = url
        self.status.config(text="Link set on the image. Right-click it to change or remove the link.",
                           foreground="green")

    def body_from_text(self, text: tk.Text) -> str:
        """Read a body box back into the stored form, with images as [[img:id]] markers."""
        out: list[str] = []
        for kind, value, _index in text.dump("1.0", "end-1c", text=True, image=True):
            if kind == "text":
                out.append(value)
            elif kind == "image" and value in self.photo_ids:
                link = self.photo_links.get(value, "")
                out.append(f"[[img:{self.photo_ids[value]}" + (f"|{link}" if link else "") + "]]")
        return "".join(out)

    def current_template(self) -> Template:
        return self.templates[self.template_box.get()]

    def on_template_change(self) -> None:
        t = self.current_template()
        self._set_entry(self.to_entry, "")  # the client's address is typed for every email
        self._set_entry(self.cc_entry, "; ".join(t.cc))

        for child in self.fields_frame.winfo_children():
            child.destroy()
        self.fill_entries.clear()
        for row, name in enumerate(placeholders(t.subject, t.body)):
            ttk.Label(self.fields_frame, text=name.replace("_", " ").capitalize()).grid(
                row=row, column=0, sticky="w", padx=4, pady=2)
            entry = ttk.Entry(self.fields_frame)
            entry.grid(row=row, column=1, sticky="ew", padx=4, pady=2)
            self.fill_entries[name] = entry

        self.fill_template()

    @staticmethod
    def _set_entry(entry: ttk.Entry, text: str) -> None:
        entry.delete(0, "end")
        entry.insert(0, text)

    def fill_template(self) -> None:
        t = self.current_template()
        values = {name: e.get().strip() for name, e in self.fill_entries.items()}
        missing = [name for name, v in values.items() if not v]
        subject = render(t.subject, values)
        company = self.company_entry.get().strip()
        if company:
            subject = replace_please_fill_in(subject, company)
        self._set_entry(self.subject_entry, subject)
        self.set_body(self.body_text, render(t.body, values))
        self.status.config(
            text=f"Still empty: {', '.join(missing)}" if missing else "Template filled. Review, then save.",
            foreground="firebrick" if missing else "gray")

    def _checked_message(self):
        """Return (to, cc, subject, body) or None if something is wrong (and shows why)."""
        to = parse_addresses(self.to_entry.get())
        cc = parse_addresses(self.cc_entry.get())
        subject = self.subject_entry.get().strip()
        body = self.body_from_text(self.body_text).strip()

        problems = []
        if not to:
            problems.append("Add at least one recipient in To.")
        bad = invalid_addresses(to + cc)
        if bad:
            problems.append("These addresses look wrong: " + ", ".join(bad))
        if not subject:
            problems.append("Subject is empty.")
        if not body:
            problems.append("Body is empty.")
        if invalid_images(body):
            problems.append("Some pasted images are missing from the template_images folder. Paste them again.")
        if self.current_template().require_attachment and not self.attachments:
            problems.append("This template needs at least one attachment. Click Add attachment first.")
        missing_files = [p for p in self.attachments if not Path(p).exists()]
        if missing_files:
            problems.append("These attachments can no longer be found: " + ", ".join(Path(p).name for p in missing_files))
        if problems:
            messagebox.showwarning("Check the email", "\n".join(problems), parent=self)
            return None
        return to, cc, subject, body

    def save_as_draft(self) -> None:
        """Saves the email to the Drafts folder of the draft account. Nothing is sent."""
        company = self.company_entry.get().strip()
        if self.current_template().name == SUBSCRIPTION_TEMPLATE and self.subscription_status(company) == "active":
            messagebox.showerror("Active mail subscription",
                                 f"{company} has an active mail subscription. {SUBSCRIPTION_TEMPLATE} can't be used "
                                 f"for them, because they need the scanned mail as an attachment. Use the Mails "
                                 "Subscription template instead.", parent=self)
            return
        if company:
            self._set_entry(self.subject_entry,
                            replace_please_fill_in(self.subject_entry.get(), company))
        msg = self._checked_message()
        if msg is None:
            return
        to, cc, subject, body = msg
        template = self.current_template().name

        choice = "skip"
        if company and not is_duplicate_preset(self.presets, company, template):
            choice = self.ask_save_preset(company, template)
            if choice == "cancel":
                return

        html_body, image_ids = body_to_html(body)
        images = [(i, image_path(i)) for i in image_ids]
        try:
            save_draft(to, cc, subject, html_body, images, list(self.attachments))
        except Exception as exc:  # show any Outlook/COM failure instead of crashing
            messagebox.showerror("Not saved", str(exc), parent=self)
            return

        try:
            log_draft(company, template, to, cc, subject)
        except OSError as exc:
            messagebox.showwarning("Log not written", f"The draft was saved, but the log failed:\n{exc}", parent=self)

        if choice == "save":
            self.presets.append(Preset(company, template, "; ".join(to)))
            save_presets(PRESETS_PATH, self.presets)
            self.log_change("preset_created", company=company, template=template, email="; ".join(to))
        self.email_done(f"Saved to Drafts: {subject}")

    # ---- Create template -------------------------------------------------

    def _build_create(self) -> None:
        frm = self.create_frame
        pad = {"padx": 8, "pady": 4}
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Template name").grid(row=0, column=0, sticky="w", **pad)
        self.new_name = ttk.Entry(frm)
        self.new_name.grid(row=0, column=1, sticky="ew", **pad)

        ttk.Label(frm, text="Default To (optional)").grid(row=1, column=0, sticky="w", **pad)
        self.new_to = ttk.Entry(frm)
        self.new_to.grid(row=1, column=1, sticky="ew", **pad)

        ttk.Label(frm, text="Default CC (separate with ;)").grid(row=2, column=0, sticky="w", **pad)
        self.new_cc = ttk.Entry(frm)
        self.new_cc.grid(row=2, column=1, sticky="ew", **pad)

        ttk.Label(frm, text="Subject").grid(row=3, column=0, sticky="w", **pad)
        self.new_subject = ttk.Entry(frm)
        self.new_subject.grid(row=3, column=1, sticky="ew", **pad)

        ttk.Label(frm, text="Body").grid(row=4, column=0, sticky="nw", **pad)
        self.new_body = tk.Text(frm, height=14, wrap="word")
        self.new_body.grid(row=4, column=1, sticky="nsew", **pad)
        self.new_body.bind("<Control-v>", self.paste_into)
        self.new_body.bind("<Button-3>", self.image_menu)
        frm.rowconfigure(4, weight=1)

        ttk.Label(frm, foreground="gray",
                  text="Use {words} for blanks, e.g. {client_name} or {due_date}. "
                       "Each one gets its own box when you use the template. "
                       "Press Ctrl+V to paste an image (copy it first, e.g. a screenshot). "
                       "Right-click an image to make it a link.",
                  wraplength=600).grid(row=5, column=0, columnspan=2, sticky="w", **pad)

        self.new_require = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="Require at least one attachment when using this template",
                        variable=self.new_require).grid(row=6, column=1, sticky="w", **pad)

        btns = ttk.Frame(frm)
        btns.grid(row=7, column=0, columnspan=2, sticky="ew", **pad)
        ttk.Button(btns, text="< Back", command=lambda: self.show(self.home)).pack(side="left", padx=4)
        ttk.Button(btns, text="Save template", command=self.save_new_template).pack(side="right", padx=4)
        ttk.Button(btns, text="Add image from file",
                   command=lambda: self.add_image_from_file(self.new_body)).pack(side="right", padx=4)

    def add_presets_for_new_template(self, template: str) -> int:
        """Gives every company that already has presets a preset for this new template.

        The new preset uses the company's email from its existing presets. Returns how many were added.
        """
        names: dict[str, str] = {}   # casefold company -> company name as first saved
        emails: dict[str, str] = {}  # casefold company -> an email from its presets
        for p in self.presets:
            key = p.company_name.strip().casefold()
            names.setdefault(key, p.company_name)
            if p.email and key not in emails:
                emails[key] = p.email

        added = 0
        for key, company in names.items():
            if is_duplicate_preset(self.presets, company, template):
                continue
            self.presets.append(Preset(company, template, emails.get(key, "")))
            added += 1
        if added:
            save_presets(PRESETS_PATH, self.presets)
            self.log_change("template_backfilled", template=template, companies_added=added)
        return added

    def save_new_template(self) -> None:
        name = self.new_name.get().strip()
        to = parse_addresses(self.new_to.get())
        cc = parse_addresses(self.new_cc.get())
        subject = self.new_subject.get().strip()
        body = self.body_from_text(self.new_body).strip()

        problems = []
        if not name:
            problems.append("Give the template a name.")
        elif name in self.templates and name != self.editing_name:
            problems.append("A template with this name already exists.")
        bad = invalid_addresses(to + cc)
        if bad:
            problems.append("These addresses look wrong: " + ", ".join(bad))
        if not subject:
            problems.append("Subject is empty.")
        if not body:
            problems.append("Body is empty.")
        if problems:
            messagebox.showwarning("Check the template", "\n".join(problems), parent=self)
            return

        new = Template(name=name, subject=subject, body=body, to=to, cc=cc,
                       require_attachment=self.new_require.get())
        if self.editing_name:
            updated = [new if t.name == self.editing_name else t for t in self.templates.values()]
            verb = "updated"
        else:
            updated = list(self.templates.values()) + [new]
            verb = "saved"
        save_templates(self.templates_path, updated)
        self.log_change("template_saved", name=name, action_type=verb)
        if self.editing_name and name != self.editing_name:
            # keep presets pointing at the renamed template
            renamed_count = sum(1 for p in self.presets if p.template_name == self.editing_name)
            for p in self.presets:
                if p.template_name == self.editing_name:
                    p.template_name = name
            save_presets(PRESETS_PATH, self.presets)
            self.log_change("preset_template_renamed", old=self.editing_name, new=name, count=renamed_count)
        self.load_templates()
        added = 0 if self.editing_name else self.add_presets_for_new_template(name)
        self.editing_name = None
        self.title("Client Reminders")
        self.clear_create_form()
        message = f'Template "{name}" {verb}.'
        if added:
            message += f"\n\nAdded a preset for it to {added} company(ies) that already have presets."
        messagebox.showinfo("Saved", message, parent=self)
        self.show(self.home)


if __name__ == "__main__":
    seed_data_folder()  # installed copy: first run only, copies the shipped data
    ReminderApp().mainloop()
