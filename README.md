# Client Email Reminders

A desktop window for drafting pre-written reminder emails for clients. It saves each email as a **draft in
Outlook classic**, so you check it and press Send yourself. Nothing is sent by the app. Presets and profiles
remember each company's template and point of contact (POC), so a reminder takes a few clicks.

**Before you start:** Outlook classic (the desktop app) must be installed and signed in to
**mail@acme-cpa.com.sg**. The new Outlook does not work with this app.

---

## For employees: installing the program

1. Go to the **Releases** page of this repository and open the latest version.
2. Under **Assets**, download `ClientRemindersSetup.exe`.
3. Run it and follow the steps. No administrator rights are needed.
4. Start **Client Email Reminders** from the Start menu or the desktop shortcut.

**Updating:** download the newest `ClientRemindersSetup.exe` and run it. The new version replaces the old one,
and your data is kept. The first time you open it, it also adds anything new since you last updated, such
as new templates, companies or presets, without changing anything you've already set yourself (your POC
choices, subscription answers, and so on). If it added anything, it tells you how many items on first open.

**Uninstalling:** open **Uninstall Client Email Reminders** from the Start menu, or use Installed apps in
Windows settings. This removes the program only. Your data is kept in `%LOCALAPPDATA%\ClientRemindersData`.

**Windows warns "Unknown publisher" or blocks the file:** the installer is not yet digitally signed. Ask
your IT team to allow it, or to sign it.

**Where things are:**
- The program: `%LOCALAPPDATA%\Programs\ClientReminders`
- Your data: `%LOCALAPPDATA%\ClientRemindersData`

---

## Using the app

### Start screen
- A **Guide** button at the top explains each part of the app.
- **Templates:** Create, Use, Edit and Delete templates.
- **Presets:** Use preset, Create preset, Profiles, Edit preset and Delete preset.

### Templates
- **Create / Edit template:** name, optional default To, default CC, subject and body. Use `{word}` for a blank,
  for example `{client_name}`. Each blank gets its own box when you use the template.
- **Images:** press Ctrl+V to paste an image (copy it first, for example a screenshot), or use
  **Add image from file**.
- **Linked images:** right-click an image and choose **Set link...** to make it open a web address
  (`https://` or `mailto:`) when clicked in the email.
- **Attachments:** a template can require at least one attachment. Tick "Require at least one attachment"
  when creating or editing it. The Urgent Letter Received and Mails Subscription templates do this.

### Presets and profiles
- A **preset** is a company plus a template. A **profile** is a company with all of its presets.
- Each profile has a **point of contact (POC)**, the address drafts are sent to. Other emails for the
  company can be kept on the profile, and any of them can be set as the POC.
- A company with no POC shows **⚠ no POC set** in the lists. You must set one before using that company
  in Use preset. The app offers to set it straight away.
- **Profiles** lists every company. Search it, sort it by name (ascending or descending), or filter it with
  the POC filter: All companies, With POC, or No POC.
- **Create preset:** enter a company name and a client email. This makes a preset for every template.

### Use preset
1. Choose a company. Use the search box, the POC filter and the sort order to find it.
2. Choose a template from the templates for that company.
3. The company's POC fills the To box. Add an attachment if the template needs one.
4. Press **Save as draft**. If the company has no preset for this template yet, you're asked whether to save one.

### Draft multiple emails
For templates that don't need an attachment (currently Mail Collection Reminder), using a preset asks whether
you want to draft for more than one client.

- **Find existing client:** search for clients, filter by POC, and sort by name. Tick **Select multiple** to
  pick several. Picks go to the **Selected clients** list, where you can remove them or press **Add selected**.
  Without Select multiple, click a client and press **Add selected**, or double-click it.
- **Enter manually:** type a company and email. You're offered a new profile for the company if it doesn't
  have one.
- **Clients to draft for:** the list of clients to draft for. Use **Remove selected** (tick **Remove multiple**
  to remove several at once) or **Clear all (keep original)**.
- **Start drafting** works through the list, one draft per client.

### Saving drafts
**Save as draft** saves the email to the Drafts folder of the account set in `reminders/outlook.py`
(`DRAFT_ACCOUNT`, currently mail@acme-cpa.com.sg). Each draft is logged to `drafts_log.csv` with the client,
template, recipients and subject.

### Change Log
The **Change Log** button on the start screen lists every data change made on this PC: templates,
presets, POCs, profile emails, names and mail subscriptions created, edited or deleted. If your presets
or profiles were set up centrally and you need IT to pick up what you've changed since, press
**Export...** and send them the file. **Clear log...** empties it once they've confirmed they have it.
**Import log...** is for IT to review and approve a log someone has sent them (see below); you shouldn't
need it day to day.

### Your data files
| File | What it is |
| --- | --- |
| `templates.json` | Your templates |
| `presets.json` | Presets (company, template, POC) |
| `profile_emails.json` | Extra emails kept on profiles |
| `person_names.json` | The name behind each email address |
| `subscriptions.json` | Mail subscription status per company |
| `change_log.jsonl` | Every data change made on this PC |
| `drafts_log.csv` / `sent_log.csv` | Log of drafts saved and emails sent |
| `template_images/` | Images used in templates |
| `companies.json`, `left_out_companies.json` | Zoho export |

The installed program keeps these in `%LOCALAPPDATA%\ClientRemindersData`. Backups (`*.before-*.json`,
`*.backup*.json`) are kept next to the data files.

---

## For developers

### Requirements
- Classic Outlook (the desktop app), installed and signed in.
- Python 3.12 from python.org (tick "Add python.exe to PATH").

### Running from source
Double-click `run.bat`. The first time, it runs `setup.bat`, which installs the packages in `requirements.txt`.
To reinstall, delete the `.setup_done` file and run `run.bat` again. Run `python app.py` to see error messages
in a terminal. When running from source, the data files live in the project folder, separate from any
installed copy.

### Publishing a release
`scripts/publish.ps1` builds the program and installer, commits and pushes the code and data, and creates a
GitHub Release with `ClientRemindersSetup.exe` attached.

1. Add a section to `CHANGELOG.md` with the version number as its heading (for example `## 1.4`), using the
   headings Added, Changed, Removed and Fixed. Write it in plain language, since employees read it.
2. Run:
   ```bash
   powershell -ExecutionPolicy Bypass -File scripts\publish.ps1 -Message "Short summary"
   ```

The version number is set automatically: 1.0 for the first release, then 0.1 higher each time. Add
`-Version` only to choose a number yourself. The release page shows the changelog section first, then a
**For developers** section with the commit list and a link to the full comparison.

Requires Inno Setup 6 and the GitHub CLI (`gh`) installed in their default locations. The publish script stops
if `.env` is staged.

### Installer
- `installer/ClientReminders.iss` is the Inno Setup script. It installs for the current user only.
- Before installing, it removes versions 1.1 and 1.2, which were built without an app ID.
- `installer/app_icon.ico` is the program and installer icon.
- The data folder is never installed or removed by the installer.

### Keeping existing installs up to date
Each release ships the current `templates.json`, `presets.json`, `profile_emails.json`, `person_names.json`
and `subscriptions.json` inside it, the same files used to seed a brand new install. On every start, the app
compares this shipped data against what's already in `%LOCALAPPDATA%\ClientRemindersData`.

**Templates always end up matching this release exactly.** A template that's missing locally is added, and
one that already exists but differs (its subject, body, To, CC or attachment requirement) is replaced with
the shipped version, even if the difference is because someone edited it on their own PC rather than through
the Change Log. The first time this happens, a message names which templates were replaced. This keeps
template wording the same everywhere; edit a template centrally (in this project) rather than on an
installed copy, since a local edit will be overwritten by the next update.

**Everything else only adds what's missing:** a new company, a new preset for an existing company, a new
email or name. It never changes or removes anything already there, so a person's own POC choices,
subscription answers and the rest are always kept. This only adds things that are new; it does not push a
correction to one of these that already exists on their PC. For that, use their Change Log (above).

To make sure the next release carries your latest data, publish from a checkout whose `presets.json` and
related files are current, for example right after applying some Change Logs.

### Zoho company import (optional)
`scripts/export_zoho_companies.py` reads contacts from Zoho CRM (read-only) and writes `companies.json` and
`left_out_companies.json`. Set it up with `.env.example` (copy it to `.env`, which must stay private and is
never shipped). `scripts/import_left_out_companies.py` adds companies to Profiles without a POC. It does a dry
run unless you pass `--apply`, skips existing profiles, and backs up the data files first.

### Applying a salesperson's Change Log
When someone sends you a `change_log.jsonl` exported from their install (the Change Log button on their
start screen), there are two ways to apply it to this project's data.

**In the app (reviewed one by one):** open Change Log on the start screen, then **Import log...**, and
pick their file. Each change is listed, newest-aware, with its current state (pending, approved, resolved,
or skipped with a reason):
- **Approve selected** applies the highlighted lines. Select several with the usual click, Shift-click
  and Ctrl-click.
- **Approve all** applies every pending change, except template changes, and always asks you to confirm
  first.
- **Resolve selected** is for template changes only (see below). It marks them as handled without
  changing any data.

**From the command line (all at once):**
```bash
python scripts/apply_change_log.py change_log.jsonl
```
It's a dry run by default; add `--apply` to make the change. Both ways use the same rules, so they never
disagree about what a change means.

Either way, template content changes (a template's subject or body) are never applied automatically.
Only the fact that a template changed is in the log, not its content, since the content can include
pasted images that only exist on that person's PC. Copy those over by hand in the app, then use
**Resolve selected** (in the app) to mark that entry as done.

### Notes
- Outlook may show a security prompt when a program saves drafts. Choose Allow.
- The body replaces the text in the draft, so the Outlook signature is not added automatically. Put your
  sign-off in the template.
- Test with your own address in To before using real clients.
- More detail on why things were built the way they are is in `DEVLOG.md`.
