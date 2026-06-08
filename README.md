# Land Survey Baseline Dashboard

Live field-monitoring dashboard for the **Land Survey Baseline — Sargodha District** study.
Built and maintained by **Research Solutions (M&A Research Solutions LLC)** · [www.rs.org.pk](https://www.rs.org.pk)

🔗 **Live site:** https://landsurvey.rs.org.pk
🔒 **Access:** password protected (`PULSE@2026`)

---

## What it shows

| Tab | Contents |
|-----|----------|
| 📊 **Overview** | Completed interviews, district target, urban/rural progress, daily submissions, tehsil coverage |
| 👩 **Respondent Profile** | Respondent type, marital status, smartphone access, household decision-maker |
| 🌾 **Land & Inheritance** | Land ownership, landholder deceased, inheritance, co-ownership |
| ⚖️ **Empowerment** | Household decision-making, mobility, relationship quality with male head |
| 🛠️ **Fieldwork & Quality** | Enumerator productivity, interview duration, consent funnel |
| 🗺️ **Mouza Completion** | Live searchable/sortable tracker — every mouza vs target, urban & rural completed |

The dashboard refreshes **daily** from the latest field data.

---

## How it works

```
Raw data (local, confidential)            Generated & deployed
──────────────────────────────            ─────────────────────
Sargodha - Land Survey Baseline.dta  ─┐
target_file.xlsx                      ─┼──►  build_dashboard.py  ──►  index.html  ──► GitHub Pages
prefill_data_PULSE.xlsx               ─┘                                              landsurvey.rs.org.pk
```

- **`build_dashboard.py`** — reads the raw files, computes every statistic
  (including the Mouza Completion tracker), and renders `index.html` from
  `template_dashboard.html`.
- **`index.html`** — fully self-contained; the only file served publicly.
  Contains aggregated statistics only — **no respondent PII**.
- **Raw data files are git-ignored** and never leave the local machine.

### Urban / Rural classification
Completed interviews are tagged urban or rural by joining on `person_id`
to the PULSE sampling tracker (`track_cat`: 3 = city/urban, 1 & 2 = village/rural).
Targets come from `target_file.xlsx` (296 mouzas across 6 tehsils).

---

## Daily update

**Manual:**
```powershell
powershell -ExecutionPolicy Bypass -File .\update_dashboard.ps1
```

**Build only (no push):**
```powershell
powershell -ExecutionPolicy Bypass -File .\update_dashboard.ps1 -NoPush
```

**Schedule (run once, elevated PowerShell):**
```powershell
$action  = New-ScheduledTaskAction -Execute "powershell.exe" `
           -Argument "-ExecutionPolicy Bypass -File `"$PWD\update_dashboard.ps1`""
$trigger = New-ScheduledTaskTrigger -Daily -At 7:00am
Register-ScheduledTask -TaskName "LandSurveyDashboardUpdate" `
           -Action $action -Trigger $trigger -Description "Daily Land Survey Dashboard refresh"
```

---

## Requirements
- Python 3 with `pandas`, `openpyxl`, `pyreadstat` (auto-installed by the update script)
- Git (for pushing) + Node optional (only used for QA)

---

© Research Solutions (M&A Research Solutions LLC) — Confidential. For authorised programme staff only.
