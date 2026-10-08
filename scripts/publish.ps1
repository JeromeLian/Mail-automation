<#
Publishes a new version of Client Reminders in one step:
  1. Sets the version in the installer.
  2. Rebuilds the program and the installer (installer_output\ClientRemindersSetup.exe).
  3. Commits the code and data changes and pushes them to GitHub.
  4. Creates a GitHub Release (v<version>) with the installer attached.

Before running: update CHANGELOG.md and README.md for this version.
Never commits .env (the Zoho credentials). The script stops if .env is staged.

Usage, from the client-reminders folder:
  powershell -ExecutionPolicy Bypass -File scripts\publish.ps1 -Message "Short summary of changes"

The version is set automatically: 1.0 for the first release, then 0.1 higher each time. Add -Version
only to choose a number yourself.
#>
param(
    [string]$Version = "",
    [Parameter(Mandatory = $true)][string]$Message,
    [string]$Notes = ""
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$gh = "C:\Program Files\GitHub CLI\gh.exe"
if (-not (Test-Path $gh)) { throw "GitHub CLI not found at $gh" }
$iscc = @("$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe", "C:\Program Files (x86)\Inno Setup 6\ISCC.exe") |
    Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw "Inno Setup not found. Install it first." }

# Version: if not given, add 0.1 to the latest GitHub release (the first release is 1.0)
if (-not $Version) {
    $latest = & $gh release list --limit 1 --json tagName -q ".[0].tagName" 2>$null
    if ($latest -match '^v(\d+)\.(\d+)') {
        $Version = "$($Matches[1]).$([int]$Matches[2] + 1)"
    } else {
        $Version = "1.0"
    }
}
if ($Version -notmatch '^\d+\.\d+(\.\d+)?$') { throw "Version must look like 1.1 or 1.1.2" }
Write-Host "Publishing version $Version"

# 1. Version number in the installer and in the app itself (shown on the start screen)
$iss = "installer\ClientReminders.iss"
(Get-Content $iss -Raw) -replace 'AppVersion=[^\r\n]+', "AppVersion=$Version" | Set-Content $iss -Encoding UTF8
$versionFile = "reminders\version.py"
(Get-Content $versionFile -Raw) -replace 'VERSION = "[^"]+"', "VERSION = `"$Version`"" | Set-Content $versionFile -Encoding UTF8
Write-Host "Version set to $Version"

# 2. Copy the shipped data into a staging folder (the .env file is never copied)
Remove-Item -Recurse -Force build_seed, dist, build -ErrorAction SilentlyContinue
New-Item -ItemType Directory build_seed | Out-Null
foreach ($file in "templates.json", "presets.json", "profile_emails.json", "companies.json",
                  "left_out_companies.json", "drafts_log.csv", "sent_log.csv", "person_names.json", "subscriptions.json") {
    if (Test-Path $file) { Copy-Item $file build_seed\ }
}
Copy-Item template_images build_seed\template_images -Recurse

# 3. Build the program, then the installer
Write-Host "Building the program..."
cmd /c "python -m PyInstaller --noconfirm --windowed --name ClientEmailReminders --icon installer\app_icon.ico --add-data build_seed;seed app.py > build_log.txt 2>&1"
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed. See build_log.txt" }
Write-Host "Building the installer..."
& $iscc installer\ClientReminders.iss > installer_build.txt 2>&1
if ($LASTEXITCODE -ne 0) { throw "Installer build failed. See installer_build.txt" }
$setup = "installer_output\ClientRemindersSetup.exe"
if (-not (Test-Path $setup)) { throw "Installer not found at $setup" }

# 4. Commit and push the code and data (stop if .env would be included)
git add -A
$staged = @(git diff --cached --name-only)
if ($staged -contains ".env") { throw ".env is staged. Unstage it with: git reset .env" }
if ($staged.Count -gt 0) {
    git commit -m "Release ${Version}: $Message" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
    if ($LASTEXITCODE -ne 0) { throw "Commit failed" }
} else {
    Write-Host "No code or data changes to commit."
}
git push origin main
if ($LASTEXITCODE -ne 0) { throw "Push failed" }

# 5. Create the release with the installer attached
$tag = "v$Version"
if ($Notes) {
    $releaseNotes = $Notes
} else {
    # The section for this version in CHANGELOG.md (heading "## <version>"), then the commits since the last release
    $changelog = Get-Content "CHANGELOG.md" -Raw
    $section = [regex]::Match($changelog, "(?ms)^## $([regex]::Escape($Version))(?=\s|$).*?(?=^## |\z)").Value.Trim()
    $lastTag = & $gh release list --limit 1 --json tagName -q ".[0].tagName" 2>$null
    $parts = @()
    if ($section) {
        $parts += $section
    } else {
        $parts += "Changes: $Message"
        Write-Host "No '## $Version' section in CHANGELOG.md. The release notes use the commit message instead."
    }
    if ($lastTag) {
        git fetch --tags -q  # the previous release's tag must exist on this PC to list its commits
        $commits = @(git log "$lastTag..HEAD" --oneline)
        $repoUrl = (& $gh repo view --json url -q .url).Trim()
        $developer = "## For developers`n`nAll changes since ${lastTag}: $repoUrl/compare/$lastTag...$tag"
        if ($commits.Count -gt 0) { $developer += "`n`nCommits:`n" + ($commits -join "`n") }
        $parts += $developer
    }
    $releaseNotes = $parts -join "`n`n"
}
Write-Host "Creating release $tag..."
# Written to a file rather than passed inline: quote characters in the notes (for example around a
# phrase quoted in the changelog) can otherwise be split into extra arguments by the native command line.
$notesFile = New-TemporaryFile
Set-Content -Path $notesFile -Value $releaseNotes -Encoding UTF8
& $gh release create $tag $setup --title "Client Reminders $Version" --notes-file $notesFile
Remove-Item $notesFile -ErrorAction SilentlyContinue
if ($LASTEXITCODE -ne 0) { throw "Release failed" }

Write-Host ""
Write-Host "Done. Release $tag is published with the installer attached."
