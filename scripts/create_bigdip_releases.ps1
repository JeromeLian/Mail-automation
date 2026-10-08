<#
Creates a GitHub Release for each version listed below on the current repo, using the matching
"## <version>" section of CHANGELOG.md as the release notes and attaching
installer_output\ClientRemindersSetup.exe to every one of them.

All releases point at the current commit (BigDip's git history was squashed to a single commit for
the presentation, so there's no older commit to tag each version against individually).

Versions skipped on purpose: 1.11, 1.12, 1.14 have no "## <version>" section in CHANGELOG.md (their
original commit messages referenced real client data that isn't carried into this sanitized repo), so
there is nothing safe to use as release notes for them.

Run this from a clean clone of JeromeLian/BigDip, authenticated as JeromeLian:
  gh repo clone JeromeLian/BigDip C:\temp\bigdip-release
  cd C:\temp\bigdip-release
  powershell -ExecutionPolicy Bypass -File scripts\create_bigdip_releases.ps1
#>
param(
    [string[]]$Versions = @("1.1","1.2","1.3","1.4","1.5","1.6","1.7","1.8","1.9","1.10","1.13","1.15","1.16")
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$setup = "installer_output\ClientRemindersSetup.exe"
if (-not (Test-Path $setup)) { throw "Installer not found at $setup. Make sure you're on the latest main." }

$changelog = Get-Content "CHANGELOG.md" -Raw

# Fetched once, up front: calling "gh release view" per tag would error on every tag that doesn't
# exist yet, and PowerShell treats a native command's stderr output as a terminating error under
# $ErrorActionPreference = "Stop", which would abort the whole script on the very first new release.
$existingTags = @(& gh release list --json tagName -q ".[].tagName" 2>$null)

foreach ($version in $Versions) {
    $tag = "v$version"

    if ($existingTags -contains $tag) {
        Write-Host "Release $tag already exists, skipping."
        continue
    }

    $section = [regex]::Match($changelog, "(?ms)^## $([regex]::Escape($version))(?=\s|$).*?(?=^## |\z)").Value.Trim()
    if (-not $section) {
        Write-Host "No '## $version' section in CHANGELOG.md, skipping $tag."
        continue
    }

    if (-not (git tag --list $tag)) {
        git tag $tag
    }
    git push origin $tag
    if ($LASTEXITCODE -ne 0) { throw "Failed to push tag $tag" }

    $notesFile = New-TemporaryFile
    Set-Content -Path $notesFile -Value $section -Encoding UTF8
    Write-Host "Creating release $tag..."
    & gh release create $tag $setup --title "Client Reminders $version" --notes-file $notesFile
    Remove-Item $notesFile -ErrorAction SilentlyContinue
    if ($LASTEXITCODE -ne 0) { throw "Release $tag failed" }
}

Write-Host ""
Write-Host "Done."
