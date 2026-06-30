<#
============================================================================
 Land Survey Baseline Dashboard - Daily Update Script
 Research Solutions (M&A Research Solutions LLC) | www.rs.org.pk
============================================================================
 What it does:
   1. (Optional) Re-exports the latest data files into this folder.
   2. Rebuilds index.html from the raw data via build_dashboard.py
      - Baseline tabs (Overview, Profile, Land, Empowerment, Vignette, Mouza)
      - Intervention tab (from "Female - Land Survey - Enumerator Script.dta";
        eligible pool = baseline-completed households, tracked by mouza)
   3. Commits and pushes the refreshed dashboard to GitHub, which
      auto-deploys to landsurvey.rs.org.pk

 NOTE: All raw data files (*.dta / *.csv / *.xlsx) are git-ignored. Only the
 rendered index.html + build sources are pushed - no respondent PII leaves
 this machine.

 Usage (manual):
   powershell -ExecutionPolicy Bypass -File .\update_dashboard.ps1

 Usage (scheduled): see "Schedule daily" section at the bottom of this file
   or run setup_schedule.ps1 once to register a Windows scheduled task.
============================================================================
#>

param(
    [switch]$NoPush  # build only, do not commit/push
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptDir

$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
function Log($msg) { Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] $msg" }

Log "============================================================"
Log " Land Survey Baseline Dashboard - Daily Update"
Log " $stamp"
Log "============================================================"

# ---------------------------------------------------------------------------
# 1. Resolve Python
# ---------------------------------------------------------------------------
$python = $null
foreach ($cmd in @("python", "py", "python3")) {
    $p = Get-Command $cmd -ErrorAction SilentlyContinue
    if ($p) { $python = $p.Source; break }
}
if (-not $python) { throw "Python was not found on PATH. Install Python 3 and retry." }
Log "Using Python: $python"

# ---------------------------------------------------------------------------
# 2. Ensure required packages
# ---------------------------------------------------------------------------
Log "Checking Python dependencies (pandas, openpyxl, pyreadstat)..."
& $python -c "import pandas, openpyxl, pyreadstat" 2>$null
if ($LASTEXITCODE -ne 0) {
    Log "Installing missing dependencies..."
    & $python -m pip install --quiet pandas openpyxl pyreadstat
}

# ---------------------------------------------------------------------------
# 3. Rebuild the dashboard
# ---------------------------------------------------------------------------
Log "Rebuilding dashboard (build_dashboard.py)..."
$env:PYTHONUTF8 = "1"
& $python build_dashboard.py
if ($LASTEXITCODE -ne 0) { throw "build_dashboard.py failed." }
Log "Dashboard rebuilt: index.html"

if ($NoPush) { Log "NoPush set - skipping git. Done."; return }

# ---------------------------------------------------------------------------
# 4. Commit & push to GitHub
# ---------------------------------------------------------------------------
$git = (Get-Command git -ErrorAction SilentlyContinue)
if (-not $git) { throw "git was not found on PATH. Install Git and retry." }

# stage only dashboard outputs/sources (never the confidential raw data)
git add index.html template_dashboard.html build_dashboard.py update_dashboard.ps1 README.md .gitignore 2>$null

$pending = git status --porcelain
if ([string]::IsNullOrWhiteSpace($pending)) {
    Log "No changes to commit. Dashboard already up to date."
} else {
    git commit -m "Daily data refresh - $stamp" | Out-Null
    Log "Committed changes."
    git push origin HEAD
    if ($LASTEXITCODE -ne 0) { throw "git push failed. Check credentials / network." }
    Log "Pushed to GitHub. Live site will update shortly."
}

Log "Done."

<#
============================================================================
 SCHEDULE DAILY (run once, in an elevated PowerShell):
----------------------------------------------------------------------------
 $action  = New-ScheduledTaskAction -Execute "powershell.exe" `
            -Argument "-ExecutionPolicy Bypass -File `"$PSScriptRoot\update_dashboard.ps1`""
 $trigger = New-ScheduledTaskTrigger -Daily -At 7:00am
 Register-ScheduledTask -TaskName "LandSurveyDashboardUpdate" `
            -Action $action -Trigger $trigger -Description "Daily refresh of Land Survey Baseline Dashboard"
============================================================================
#>
