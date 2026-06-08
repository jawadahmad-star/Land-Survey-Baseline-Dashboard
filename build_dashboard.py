# -*- coding: utf-8 -*-
"""
Land Survey Baseline Dashboard — Data Builder
=============================================
Research Solutions (M&A Research Solutions LLC) | www.rs.org.pk

Reads the raw survey + sampling files, computes every statistic the dashboard
needs (including the Mouza Completion tracker), and renders a fully
self-contained `index.html` from `template_dashboard.html`.

Run:  python build_dashboard.py
Daily automation is handled by update_dashboard.ps1
"""
import json
import sys
import datetime as dt
from pathlib import Path

import pandas as pd
import pyreadstat

# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------
BASE = Path(__file__).resolve().parent
DTA_FILE      = BASE / "Sargodha - Land Survey Baseline.dta"
TARGET_FILE   = BASE / "target_file.xlsx"
PREFILL_FILE  = BASE / "prefill_data_PULSE.xlsx"
TEMPLATE_FILE = BASE / "template_dashboard.html"
OUTPUT_FILE   = BASE / "index.html"

COMPLETE_STATUS = 1          # status_survey == 1  => completed interview
URBAN_CATS = {3}             # track_cat: 3 = city  (urban)
RURAL_CATS = {1, 2}          # track_cat: 1,2 = village (rural)


def vc(series, labels=None, dropna=True):
    """value_counts as ordered list of {label, value} using a label map."""
    counts = series.value_counts(dropna=dropna)
    out = []
    for k, v in counts.items():
        if pd.isna(k):
            name = "No response"
        elif labels and k in labels:
            name = labels[k]
        else:
            name = str(k)
        out.append({"label": name, "value": int(v)})
    return out


def mean_scale(series):
    s = pd.to_numeric(series, errors="coerce")
    s = s[s <= 5]                     # guard against 666/999 codes
    return round(float(s.mean()), 2) if len(s.dropna()) else 0.0


def main():
    # ------------------------------------------------------------------
    # Load
    # ------------------------------------------------------------------
    df, meta = pyreadstat.read_dta(str(DTA_FILE))
    vlabels = meta.variable_value_labels
    target = pd.read_excel(TARGET_FILE)
    prefill = pd.read_excel(PREFILL_FILE)

    # ------------------------------------------------------------------
    # Completed interviews + urban/rural tagging via prefill track_cat
    # ------------------------------------------------------------------
    comp = df[df["status_survey"] == COMPLETE_STATUS].copy()
    comp["pid"] = comp["person_id"].astype(str).str.replace(r"\.0$", "", regex=True)
    prefill["pid"] = prefill["person_id"].astype(str)
    pmap = prefill[["pid", "track_cat"]].drop_duplicates("pid")
    comp = comp.merge(pmap, on="pid", how="left")
    comp["is_urban"] = comp["track_cat"].isin(URBAN_CATS)
    comp["is_rural"] = comp["track_cat"].isin(RURAL_CATS)
    # records that did not match prefill: classify by mauza dominant later;
    # default unmatched to rural (village) since 95%+ of frame is rural
    unmatched = comp["track_cat"].isna()
    comp.loc[unmatched, "is_rural"] = True

    n_complete = len(comp)
    n_submissions = len(df)

    # ------------------------------------------------------------------
    # Daily submissions (use starttime date)
    # ------------------------------------------------------------------
    daily = {}
    if "starttime" in df.columns:
        st = pd.to_datetime(df["starttime"], errors="coerce")
        for d in st.dropna().dt.date:
            key = d.isoformat()
            daily[key] = daily.get(key, 0) + 1
    daily_list = [{"date": k, "count": v} for k, v in sorted(daily.items())]

    # ------------------------------------------------------------------
    # Targets (district-wide)
    # ------------------------------------------------------------------
    total_target = int(target["target_final"].sum())
    urban_target = int(target["urban_target"].sum())
    rural_target = int(target["rural_target"].sum())
    n_target_mauzas = int(target["Mauza_sample"].nunique())
    n_tehsils = int(target["Tehsil_sample"].nunique())

    urban_done = int(comp["is_urban"].sum())
    rural_done = int(comp["is_rural"].sum())

    # ------------------------------------------------------------------
    # Mouza completion table
    # ------------------------------------------------------------------
    done_by_mauza = comp.groupby("mauza").agg(
        urban_done=("is_urban", "sum"),
        rural_done=("is_rural", "sum"),
        total_done=("is_urban", "size"),
    )
    rows = []
    for _, r in target.iterrows():
        mauza = r["Mauza_sample"]
        ud = int(done_by_mauza["urban_done"].get(mauza, 0)) if mauza in done_by_mauza.index else 0
        rd = int(done_by_mauza["rural_done"].get(mauza, 0)) if mauza in done_by_mauza.index else 0
        td = int(done_by_mauza["total_done"].get(mauza, 0)) if mauza in done_by_mauza.index else 0
        tt = int(r["target_final"])
        ut = int(r["urban_target"])
        rt = int(r["rural_target"])
        pct = round(100 * td / tt, 1) if tt else 0.0
        if td == 0:
            status = "Not Started"
        elif td >= tt:
            status = "Completed"
        else:
            status = "In Progress"
        rows.append({
            "mauza": mauza,
            "tehsil": r["Tehsil_sample"],
            "urban_done": ud, "urban_target": ut,
            "rural_done": rd, "rural_target": rt,
            "total_done": td, "total_target": tt,
            "pct": pct, "status": status,
        })
    # sort: in-progress/completed first (by completion desc), then not started
    rows.sort(key=lambda x: (x["total_done"] == 0, -x["total_done"], -x["pct"]))

    mauzas_started = sum(1 for x in rows if x["total_done"] > 0)
    mauzas_completed = sum(1 for x in rows if x["status"] == "Completed")

    # ------------------------------------------------------------------
    # Panel distributions (completed interviews only)
    # ------------------------------------------------------------------
    def lab(col):
        return vlabels.get(col)

    resp_type   = vc(comp["resp_type"],     lab("resp_type"),     dropna=False)
    marital     = vc(comp["marital_status"], lab("marital_status"), dropna=False)
    smartphone  = vc(comp["smart_phone"],   lab("smart_phone"),   dropna=False)
    tehsil_split = vc(comp["tehsil"], dropna=False)

    land_in_name = vc(comp["land_in_name"], lab("land_in_name"), dropna=False)
    father_died  = vc(comp["father_died"],  lab("father_died"),  dropna=False)
    inherit_land = vc(comp["inherit_land"], lab("inherit_land"), dropna=False)
    co_own_land  = vc(comp["co_own_land"],  lab("co_own_land"),  dropna=False)

    decision_money = vc(comp["s4_q1"], lab("s4_q1"), dropna=False)
    decision_daily = vc(comp["s4_q2"], lab("s4_q2"), dropna=False)
    decision_save  = vc(comp["s4_q3"], lab("s4_q3"), dropna=False)

    # mobility: days outside village alone (numeric text)
    mob = pd.to_numeric(comp["s4_q5"], errors="coerce")
    mobility = []
    for bucket, cond in [("0 days", mob == 0), ("1–2 days", mob.isin([1, 2])),
                         ("3–4 days", mob.isin([3, 4])), ("5+ days", mob >= 5)]:
        mobility.append({"label": bucket, "value": int(cond.sum())})

    # relationship quality means (s4_q8/9/10 on 1-4)
    rel = {
        "Shows interest in me": mean_scale(comp["s4_q8"]),
        "Respects me":          mean_scale(comp["s4_q9"]),
        "I can share problems": mean_scale(comp["s4_q10"]),
    }

    # ------------------------------------------------------------------
    # Fieldwork & quality
    # ------------------------------------------------------------------
    enum = (comp["enum_label"].value_counts().head(15)
            if "enum_label" in comp.columns else pd.Series(dtype=int))
    enumerators = [{"label": k, "value": int(v)} for k, v in enum.items()]

    dur = pd.to_numeric(comp["duration"], errors="coerce") / 60.0  # minutes
    dur_bins = []
    for label, cond in [("<30 min", dur < 30), ("30–45", (dur >= 30) & (dur < 45)),
                        ("45–60", (dur >= 45) & (dur < 60)), ("60–75", (dur >= 60) & (dur < 75)),
                        ("75+ min", dur >= 75)]:
        dur_bins.append({"label": label, "value": int(cond.sum())})
    avg_duration = round(float(dur.median()), 0) if len(dur.dropna()) else 0

    # consent funnel
    n_intro_yes = int((df["intro_consent"].isin([1, 3])).sum())
    n_consent_yes = int((df["consent"] == 1).sum())
    consent_rate = round(100 * n_consent_yes / n_submissions, 1) if n_submissions else 0

    field_days = len(daily_list)

    # ------------------------------------------------------------------
    # Assemble payload
    # ------------------------------------------------------------------
    data = {
        "meta": {
            "last_updated": dt.datetime.now().strftime("%d %b %Y, %I:%M %p"),
            "last_date": daily_list[-1]["date"] if daily_list else "",
            "n_complete": n_complete,
            "n_submissions": n_submissions,
            "total_target": total_target,
            "urban_target": urban_target,
            "rural_target": rural_target,
            "urban_done": urban_done,
            "rural_done": rural_done,
            "n_target_mauzas": n_target_mauzas,
            "n_tehsils": n_tehsils,
            "mauzas_started": mauzas_started,
            "mauzas_completed": mauzas_completed,
            "consent_rate": consent_rate,
            "field_days": field_days,
            "avg_duration": avg_duration,
            "pct_complete": round(100 * n_complete / total_target, 1) if total_target else 0,
        },
        "daily": daily_list,
        "tehsil_split": tehsil_split,
        "resp_type": resp_type,
        "marital": marital,
        "smartphone": smartphone,
        "land_in_name": land_in_name,
        "father_died": father_died,
        "inherit_land": inherit_land,
        "co_own_land": co_own_land,
        "decision_money": decision_money,
        "decision_daily": decision_daily,
        "decision_save": decision_save,
        "mobility": mobility,
        "relationship": rel,
        "enumerators": enumerators,
        "duration_bins": dur_bins,
        "consent_funnel": [
            {"label": "Submissions", "value": n_submissions},
            {"label": "Intro consent", "value": n_intro_yes},
            {"label": "Consented", "value": n_consent_yes},
            {"label": "Completed", "value": n_complete},
        ],
        "mouza_table": rows,
    }

    # ------------------------------------------------------------------
    # Render
    # ------------------------------------------------------------------
    template = TEMPLATE_FILE.read_text(encoding="utf-8")
    payload = json.dumps(data, ensure_ascii=False)
    html = template.replace("/*__DASHBOARD_DATA__*/null", payload)
    OUTPUT_FILE.write_text(html, encoding="utf-8")

    print(f"[OK] index.html generated  ({OUTPUT_FILE})")
    print(f"     Completed interviews : {n_complete}")
    print(f"     Urban / Rural done   : {urban_done} / {rural_done}")
    print(f"     District target      : {total_target}  ({urban_target} urban / {rural_target} rural)")
    print(f"     Mauzas started       : {mauzas_started} / {n_target_mauzas}")
    print(f"     Last field date      : {data['meta']['last_date']}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("[ERROR]", e, file=sys.stderr)
        raise
