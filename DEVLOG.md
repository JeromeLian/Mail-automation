# Developer log

A running record of how and why the app is built the way it is. Newest entries first.
For what changed in each release, see CHANGELOG.md.

---

## 2026-10-07: packaging, releases and the Guide

**What:** packaged the program with PyInstaller, wrapped it in an installer with Inno Setup,
added the Guide window, and wrote `scripts/publish.ps1`. Version 1.0 was released.

**Decisions and why:**
- **Unsigned installer.** No code-signing certificate yet. Antivirus may warn on first run.
- **Per-user install** (`%LOCALAPPDATA%\Programs\ClientReminders`). No administrator rights needed.
- **Data kept apart from the program** (`%LOCALAPPDATA%\ClientRemindersData`). Updates and uninstalls
  don't touch it.
- **First run copies shipped data** (`seed_data_folder` in `reminders/core.py`). It only copies files
  that don't exist yet, so existing users keep their data. Downside: they won't receive new templates
  or companies from later versions. A merge step is needed for that.
- **Releases on the same private repo.** The installer is attached to a GitHub Release rather than
  committed, to keep binaries out of history. Anyone with access to the repo can see the code and data.
  Access is through one shared account, so downloads can't be traced to a person.
- **Versions** are set by `publish.ps1`: 1.0 first, then +0.1 from the latest release.
  Release notes come from the matching `CHANGELOG.md` section plus commits since the last release.

**Known gaps:**
- The publish script needs Inno Setup and the GitHub CLI installed in their default locations.
  It doesn't check that the installer works. Test the installer before each release.
- Test records exist in the Zoho data (for example "abc", "test annex b"). Eight existing profiles are
  test records and have not been removed.

---

## 2026-10-06 to 2026-10-07: profiles, POCs and the draft workflow

**What:** profiles group each company's presets. Each profile has a point of contact (POC).
Use preset and Draft multiple emails both depend on the POC.

**Decisions and why:**
- **POC is the email stored on the company's presets.** The POC is not a separate field. Existing
  presets already held the address, so the POC is simply the address they use. This meant no data
  migration. Extra emails live in `profile_emails.json`.
- **Companies without a POC are allowed.** They appear with a warning and the app offers to set one before
  they're used. This avoids silently drafting to the wrong address.
- **Never overwrite existing profiles on import.** The Zoho import skips any company whose name matches an
  existing profile, including near-matches (ignoring "Pte Ltd" and punctuation). Those are listed for review.
- **Lists are built once per refresh.** Looking up each company's POC by scanning all presets made typing
  laggy. A single pass over the presets (`poc_index`) fixed it.
- **Clicking a client only selects it.** Multi-select uses a separate "Selected clients" list so that a
  search or filter change can't lose picks.

**Known gaps:**
- Clicks in the Find existing client list rebuild the list only when needed; a large list may still feel
  slow on older PCs.

---

## 2026-10-06: Zoho export (read-only)

**What:** `scripts/export_zoho_companies.py` reads contacts through the Zoho CRM API and writes
`companies.json` (companies with one client email, or two with one @acme-cpa.com.sg address) and
`left_out_companies.json` (everything else, with their client emails).

**Decisions and why:**
- **Read-only.** The script only makes GET requests. Zoho is never written to.
- **Credentials in `.env`**, which is excluded from git and never shipped in the installer.
- **Paging with `page_token`.** Zoho refuses numbered pages past 2,000 records.
- **Firm addresses are never used as a client email.** Staff addresses sit on many companies.

---

## 2026-10-06: initial app

**What:** Tkinter desktop app. Templates with `{placeholder}` blanks, sending through Outlook classic via
COM, pasted images, attachments, presets, and a send log.

**Decisions and why:**
- **Tkinter**, so there are no extra front-end dependencies.
- **Outlook classic COM** instead of Microsoft Graph, so no tenant permissions are needed.
  The new Outlook has no COM interface, so it can't be used.
- **Drafts instead of sending.** Sending was replaced with "Save as draft" so a person always presses Send.
- **Images stored as files and marked in the template body** (`[[img:<id>]]`), so templates stay plain text.
