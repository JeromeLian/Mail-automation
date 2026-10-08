# Changelog

## 1.16

### Added
- The Profiles search box now also matches a person's name, not just the company name — the same rule
  Use preset and Find existing client already use. When a result matches by person, their name is shown
  next to the company.

## 1.15

### Changed
- Templates now always match this release: a missing one is added, and one that already exists but
  differs (including a local edit made on that PC) is replaced with the central version. You're told
  which templates were replaced the first time it happens. Everything else (presets, POCs, profile
  emails, names, subscriptions) still only adds what's missing, as before.

## 1.13

### Changed
- Running the app from source (`run.bat`) now shows "(dev)" after the version number on the start
  screen, so a dev copy is never mistaken for the installed release.

## 1.10

### Added
- Draft multiple emails: the page now scrolls, so Enter manually no longer runs off the bottom of the window.
- Find existing client (Draft multiple emails) now offers to set the POC for a company added without one,
  the same prompt as Use preset. Adding several at once (Select multiple) asks for each one in turn.
- Clients to draft for now shows the contact's name next to their email, not just the address.

### Fixed
- A company with a saved email but no POC was wrongly treated as having no email at all when added from
  Find existing client, and sent to manual entry instead. It's now added normally and offered the POC prompt.

## 1.9

### Added
- Updating now also brings in new central data automatically: new templates, companies, presets, emails and names from the update are added to what's already on that PC. Nothing already set locally is changed or removed. A message on first open after an update says how many items were added.

## 1.8

### Added
- Change Log: every data change made in the app (templates, presets, POCs, profile emails, names, mail subscriptions) is now recorded, with an Export... button to send it on.
- Import log...: IT can load someone's exported Change Log in the app and review each change before approving it, select several at once, or approve all pending changes with a confirmation first.
- Resolve selected: template changes (which can't be auto-applied, since their content isn't in the log) are marked as handled by hand, separately from Approve.
- scripts/apply_change_log.py applies an exported Change Log from the command line, using the same rules as the in-app review.

## 1.7

### Added
- The version number is shown at the bottom of the start screen, along with a small credit line.

### Changed
- The window opens closer to the top of the screen.
- The Profiles subscription filter now has four options: Active subscription, Past subscription (not active), No subscription, and All companies. A past subscription that isn't active is no longer grouped with companies that never had one.

## 1.6

### Added
- Mail subscriptions: companies with a past "Mail Scanning" or "Mail Subscription" invoice in Zoho are now recorded on their profile.
- Using Mail Collection Reminder for a client with a recorded subscription asks whether it's currently active. Yes asks for an end date using a calendar picker. No means you won't be asked again.
- A client with an active subscription is asked to confirm it's still active each time the template is used, and is checked again automatically once the end date passes.
- A client with an active subscription can't be drafted with Mail Collection Reminder, since they need the scanned mail attached. Mails Subscription should be used instead.
- The draft multiple emails page shows a notice listing selected clients whose subscription isn't active.
- Edit subscription on the profile screen sets the status and end date by hand, including for companies with no Zoho match.
- Dates are shown and entered as dd/mm/yyyy, with a calendar picker, and invalid dates are rejected.
- The Profiles list has a subscription filter (All companies, Has a subscription, No subscription), and each filter now has a short label.

## 1.5

### Added
- Names for people: each address on a profile can have a person's name. Names were filled in from Zoho where available. You can add or edit a name with Edit name on the profile screen.
- Search by person: typing a person's name in Use preset or Find existing client finds their company.
- When a name is entered for a client in Enter manually, you choose whether they are the point of contact (POC).

### Changed
- A point of contact must have a name before it can be used.
- Enter manually takes one email address only.
- In Find existing client, the Selected clients list sits beside the search results, and Enter manually is below them.

## 1.4

### Added
- The installer removes earlier versions (1.1 and 1.2) before installing, so you don't need to uninstall them by hand. Your data is kept.
- An updated README covering installation, updating, uninstalling and the Guide.

## 1.3

### Changed
- The program is now called Client Email Reminders in the Start menu, on the desktop and in Installed apps.
- Updating from this version onward replaces the old version in place, so there are no duplicate entries.

## 1.2

### Added
- A new program and installer icon: the John Carter logo with an email symbol.
- A reminder on the start screen to sign in to mail@acme-cpa.com.sg in Outlook classic before using the app.

## 1.1

### Changed
- Test release to check the new release page layout. The app itself has not changed, so there is nothing new to install.

## 2026-10-07

### Added
- Template images: paste with Ctrl+V or add from file, and set an image to open a web address when clicked.
- Attachments on the email screen, with a per-template "require at least one attachment" option.
- Presets and profiles: create, edit, delete, and a Profiles screen grouping each company's presets.
- Points of contact (POC): a POC per company, extra emails per company, and a "no POC set" alert.
  Use preset offers to set the POC before continuing.
- Use preset: company and template search, with the POC filter, sorting and "no POC" alerts.
- Draft multiple emails for Mail Collection Reminder, with Find existing client, Enter manually,
  Select multiple, the Selected clients list, Remove multiple and Clear all (keep original).
- Zoho company export (read-only) and an import of companies left out of the export, with a dry run.
- Search, POC filter and sort on the Profiles list and the client search lists.

### Changed
- Save as draft replaces Send. Drafts are saved to the Outlook account set in `DRAFT_ACCOUNT`.
- Presets store an email (the POC), and creating a profile creates a preset for every template.
- The app opens near the top of the screen, and its height adjusts to fit.
- Lists are built once per refresh rather than once per company, which removed lag when typing.
- Clicking a client in the search list no longer adds it. Use Add selected or double-click.

### Removed
- The "Date added" sort option. It was not needed.

### Fixed
- Presets with no email now fall back to the most recent draft to that company.
- The Urgent Letter Received and Mails Subscription templates require an attachment.

## Initial version

- Outlook classic integration, templates with `{placeholder}` blanks, and sending through Outlook.
