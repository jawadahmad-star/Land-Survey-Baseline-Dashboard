# -*- coding: utf-8 -*-
"""
Mauza Touch Tracker
===================
Splits the sampling frame (target_file.xlsx) into:
  * Untouched_Mauza_NoRespondent.xlsx -> mauzas where NOT a single respondent
    has been touched in either survey sheet.
  * Touched_Mauza.xlsx -> mauzas where at least one respondent was touched,
    with urban / rural touched counts.

A mauza is "touched" if its name appears in the `mauza` column of either
survey export (narrow or WIDE). Urban/rural classification of each respondent
follows the dashboard logic: prefill track_cat 3 = urban, 1/2 = rural,
unmatched records default to rural.
"""
from pathlib import Path
import pandas as pd

BASE = Path(__file__).resolve().parent
NARROW = BASE / "Sargodha - Land Survey Baseline_.csv"
WIDE   = BASE / "Sargodha - Land Survey Baseline_WIDE.csv"
TARGET = BASE / "target_file.xlsx"
PREFILL = BASE / "prefill_data_PULSE.xlsx"

URBAN_CATS = {3}
RURAL_CATS = {1, 2}

# ---- load survey data (both sheets) ----------------------------------------
narrow = pd.read_csv(NARROW, dtype=str, low_memory=False)
wide   = pd.read_csv(WIDE, dtype=str, low_memory=False)

# combined respondent frame, de-duplicated on KEY (narrow is a subset of WIDE)
key_col = "KEY" if "KEY" in wide.columns else "instanceID"
survey = pd.concat([wide, narrow], ignore_index=True)
survey = survey.drop_duplicates(subset=[key_col])
survey = survey[survey["mauza"].notna() & (survey["mauza"].astype(str).str.strip() != "")]

touched_mauzas = set(survey["mauza"].astype(str))

# ---- urban / rural classification via prefill track_cat --------------------
prefill = pd.read_excel(PREFILL)
survey["pid"] = survey["person_id"].astype(str).str.replace(r"\.0$", "", regex=True)
prefill["pid"] = prefill["person_id"].astype(str).str.replace(r"\.0$", "", regex=True)
pmap = prefill[["pid", "track_cat"]].drop_duplicates("pid")
survey = survey.merge(pmap, on="pid", how="left")
survey["is_urban"] = survey["track_cat"].isin(URBAN_CATS)
survey["is_rural"] = survey["track_cat"].isin(RURAL_CATS)
# unmatched (no prefill) default to rural, matching dashboard
survey.loc[survey["track_cat"].isna(), "is_rural"] = True

counts = survey.groupby("mauza").agg(
    Urban_Touched=("is_urban", "sum"),
    Rural_Touched=("is_rural", "sum"),
    Total_Touched=("is_urban", "size"),
)

# completed interviews only (status_survey == 1) -> matches the dashboard's
# "Completed" definition. Touched counts above include refusals/locked/etc.
COMPLETE_STATUS = "1"
completed_counts = (survey[survey["status_survey"].astype(str) == COMPLETE_STATUS]
                    .groupby("mauza").size())

# ---- target frame -----------------------------------------------------------
target = pd.read_excel(TARGET)

untouched_rows, touched_rows = [], []
for _, r in target.iterrows():
    mauza = str(r["Mauza_sample"])
    base = {
        "Mauza": mauza,
        "Tehsil": r["Tehsil_sample"],
        "Rural_Target": int(r["rural_target"]),
        "Urban_Target": int(r["urban_target"]),
        "Target_Final": int(r["target_final"]),
    }
    if mauza in touched_mauzas:
        row = dict(base)
        row["Urban_Touched"] = int(counts.loc[mauza, "Urban_Touched"]) if mauza in counts.index else 0
        row["Rural_Touched"] = int(counts.loc[mauza, "Rural_Touched"]) if mauza in counts.index else 0
        row["Total_Touched"] = int(counts.loc[mauza, "Total_Touched"]) if mauza in counts.index else 0
        # completed interviews (status_survey == 1) drive the Status, matching the dashboard
        row["Completed_Interviews"] = int(completed_counts.get(mauza, 0))
        row["Status"] = "Completed" if row["Completed_Interviews"] >= row["Target_Final"] else "In Progress"
        touched_rows.append(row)
    else:
        untouched_rows.append(base)

untouched = pd.DataFrame(untouched_rows).sort_values(["Tehsil", "Mauza"]).reset_index(drop=True)
touched = pd.DataFrame(touched_rows).sort_values(["Tehsil", "Mauza"]).reset_index(drop=True)

untouched.to_excel(BASE / "Untouched_Mauza_NoRespondent.xlsx", index=False)
touched.to_excel(BASE / "Touched_Mauza.xlsx", index=False)

print(f"Target mauzas        : {len(target)}")
print(f"Touched mauzas       : {len(touched)}")
print(f"Untouched mauzas     : {len(untouched)}")
print(f"Respondents (deduped): {len(survey)}  "
      f"(urban {int(survey['is_urban'].sum())} / rural {int(survey['is_rural'].sum())})")
print("[OK] Untouched_Mauza_NoRespondent.xlsx + Touched_Mauza.xlsx written")
