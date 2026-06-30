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
IV_FILE       = BASE / "Female - Land Survey - Enumerator Script.dta"
TARGET_FILE   = BASE / "target_file.xlsx"
PREFILL_FILE  = BASE / "prefill_data_PULSE.xlsx"
TEMPLATE_FILE = BASE / "template_dashboard.html"
OUTPUT_FILE   = BASE / "index.html"

COMPLETE_STATUS = 1          # status_survey == 1  => completed interview
URBAN_CATS = {3}             # track_cat: 3 = city  (urban)
RURAL_CATS = {1, 2}          # track_cat: 1,2 = village (rural)

# Headline sampling target shown on the dashboard.
# Set to None to use the per-mouza sums from target_file.xlsx instead.
# (Per-mouza Mouza Completion table always uses target_file.xlsx.)
TARGET_OVERRIDE = {"total": 3750, "urban": 1250, "rural": 2500}


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


def build_intervention(comp, n_complete):
    """Compute the Intervention tab payload.

    `comp` is the de-duplicated, completed BASELINE dataframe (the eligible
    pool — only households whose baseline was completed receive the
    intervention). `n_complete` is len(comp), the total eligible households.
    """
    out = {
        "available": False,
        "meta": {}, "daily": [], "treat": [], "arm": [],
        "awareness": [], "knows_listed": [], "knows_fard": [],
        "enumerators": [], "mauza_table": [],
    }
    if not IV_FILE.exists():
        return out

    ivdf, ivmeta = pyreadstat.read_dta(str(IV_FILE))

    # Completed intervention visits, de-duplicated on hh_id (earliest start)
    ivc = ivdf[ivdf["status_survey"] == COMPLETE_STATUS].copy()
    if "hh_id" in ivc.columns:
        if "starttime" in ivc.columns:
            ivc = ivc.sort_values("starttime")
        ivc = ivc.drop_duplicates(subset="hh_id", keep="first")

    n_iv = len(ivc)
    n_iv_submissions = len(ivdf)
    eligible = int(n_complete)

    # ---- Eligible pool & intervention done, per mouza (from baseline comp) --
    # Drop records with a missing/blank mauza (a few baseline rows have no
    # mauza/tehsil recorded — they must not appear as a phantom mouza).
    comp_m = comp[comp["mauza"].notna() & (comp["mauza"].astype(str).str.strip() != "")]
    ivc_m  = ivc[ivc["mauza"].notna() & (ivc["mauza"].astype(str).str.strip() != "")]
    elig_by_mauza = comp_m.groupby("mauza").size()
    tehsil_by_mauza = (comp_m.groupby("mauza")["tehsil"]
                       .agg(lambda s: s.mode().iat[0] if len(s.mode()) else ""))
    done_by_mauza = ivc_m.groupby("mauza").size()

    mauza_rows = []
    for mauza, elig in elig_by_mauza.items():
        elig = int(elig)
        done = int(done_by_mauza.get(mauza, 0))
        pct = round(100 * done / elig, 1) if elig else 0.0
        if elig and done >= elig:
            status = "Completed"
        elif done > 0:
            status = "In Progress"
        else:
            status = "Not Started"
        mauza_rows.append({
            "mauza": mauza,
            "tehsil": tehsil_by_mauza.get(mauza, ""),
            "done": done, "eligible": elig,
            "touched": done > 0, "pct": pct, "status": status,
        })
    # started (any intervention) first, by completion desc; rest after
    mauza_rows.sort(key=lambda x: (not x["touched"], -x["done"], -x["pct"]))

    mauzas_reached = sum(1 for x in mauza_rows if x["touched"])
    mauzas_done = sum(1 for x in mauza_rows if x["status"] == "Completed")

    # ---- Daily intervention visits ----
    daily = {}
    if "starttime" in ivc.columns:
        st = pd.to_datetime(ivc["starttime"], errors="coerce")
        for d in st.dropna().dt.date:
            daily[d.isoformat()] = daily.get(d.isoformat(), 0) + 1
    daily_list = [{"date": k, "count": v} for k, v in sorted(daily.items())]
    last_date = daily_list[-1]["date"] if daily_list else ""

    # ---- Treatment assignment (T1 / T2) ----
    treat = vc(ivc["treat"].astype(str).str.strip().replace({"": "Unassigned"})) \
        if "treat" in ivc.columns else []

    # ---- Treatment arm (women only / women + men) ----
    arm_raw = (ivc["treat_arm"].astype(str).str.strip()
               .replace({"": "Not recorded", "nan": "Not recorded"})) \
        if "treat_arm" in ivc.columns else pd.Series(dtype=str)
    arm = [{"label": k, "value": int(v)} for k, v in arm_raw.value_counts().items()]

    # ---- Substantive awareness indicators (intervention questions) ----
    # Explicit, clean label maps (the .dta stores apostrophes as mojibake).
    AWARE_LAB = {1: "Knows the initiative", 2: "Never heard of it",
                 3: "Heard of it, knows nothing"}
    YESNO_LAB = {1: "Yes", 2: "No"}
    FARD_LAB  = {1: "Yes", 2: "No", 88: "Refused", 99: "Don't know"}
    awareness    = vc(ivc["p5_q5"],      AWARE_LAB, dropna=False) if "p5_q5" in ivc.columns else []
    knows_listed = vc(ivc["part3_q2_2"], YESNO_LAB, dropna=False) if "part3_q2_2" in ivc.columns else []
    knows_fard   = vc(ivc["fard_intro"], FARD_LAB,  dropna=False) if "fard_intro" in ivc.columns else []

    # ---- Enumerators ----
    if "enum_label" in ivc.columns:
        ie = ivc["enum_label"].astype(str).replace({"": "—", "nan": "—"}).value_counts().head(12)
        enumerators = [{"label": k, "value": int(v)} for k, v in ie.items()]
    else:
        enumerators = []

    dur = pd.to_numeric(ivc["duration"], errors="coerce") / 60.0 if "duration" in ivc.columns else pd.Series(dtype=float)
    avg_dur = round(float(dur.median()), 0) if len(dur.dropna()) else 0

    out.update({
        "available": True,
        "meta": {
            "n_iv": n_iv,
            "n_iv_submissions": n_iv_submissions,
            "eligible": eligible,
            "pct": round(100 * n_iv / eligible, 1) if eligible else 0.0,
            "mauzas_reached": mauzas_reached,
            "mauzas_done": mauzas_done,
            "mauzas_eligible": int(len(mauza_rows)),
            "field_days": len(daily_list),
            "last_date": last_date,
            "avg_duration": avg_dur,
            "n_t1": int(ivc["treat"].astype(str).str.strip().eq("T1").sum()) if "treat" in ivc.columns else 0,
            "n_t2": int(ivc["treat"].astype(str).str.strip().eq("T2").sum()) if "treat" in ivc.columns else 0,
        },
        "daily": daily_list,
        "treat": treat,
        "arm": arm,
        "awareness": awareness,
        "knows_listed": knows_listed,
        "knows_fard": knows_fard,
        "enumerators": enumerators,
        "mauza_table": mauza_rows,
    })
    return out


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
    # De-duplicate on hh_id: a household interview submitted/synced more than
    # once must count only ONCE, otherwise the completed total is inflated.
    # Keep the earliest submission (by starttime) for each hh_id.
    if "hh_id" in comp.columns:
        if "starttime" in comp.columns:
            comp = comp.sort_values("starttime")
        comp = comp.drop_duplicates(subset="hh_id", keep="first")
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
    # Daily COMPLETED interviews (sums to n_complete = 76)
    # ------------------------------------------------------------------
    daily = {}
    if "starttime" in comp.columns:
        st = pd.to_datetime(comp["starttime"], errors="coerce")
        for d in st.dropna().dt.date:
            key = d.isoformat()
            daily[key] = daily.get(key, 0) + 1
    daily_list = [{"date": k, "count": v} for k, v in sorted(daily.items())]

    # ------------------------------------------------------------------
    # Targets (district-wide)
    # ------------------------------------------------------------------
    if TARGET_OVERRIDE:
        total_target = int(TARGET_OVERRIDE["total"])
        urban_target = int(TARGET_OVERRIDE["urban"])
        rural_target = int(TARGET_OVERRIDE["rural"])
    else:
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
    # touched = field team reached the mauza (ANY survey record, any status:
    # completed, locked/empty, refused, respondent died, etc.). This matches
    # the Touched_Mauza tracker (build_mauza_touch.py); a mauza can be touched
    # without a single completed interview yet.
    touched_src = df[df["mauza"].notna() & (df["mauza"].astype(str).str.strip() != "")]
    touched_by_mauza = touched_src.groupby("mauza").size()
    rows = []
    for _, r in target.iterrows():
        mauza = r["Mauza_sample"]
        ud = int(done_by_mauza["urban_done"].get(mauza, 0)) if mauza in done_by_mauza.index else 0
        rd = int(done_by_mauza["rural_done"].get(mauza, 0)) if mauza in done_by_mauza.index else 0
        td = int(done_by_mauza["total_done"].get(mauza, 0)) if mauza in done_by_mauza.index else 0
        touched = bool(mauza in touched_by_mauza.index)
        tt = int(r["target_final"])
        ut = int(r["urban_target"])
        rt = int(r["rural_target"])
        pct = round(100 * td / tt, 1) if tt else 0.0
        if tt and td >= tt:
            status = "Completed"
        elif td > 0 or touched:
            # field work has begun (completed interviews and/or visits made)
            status = "In Progress"
        else:
            status = "Not Started"
        rows.append({
            "mauza": mauza,
            "tehsil": r["Tehsil_sample"],
            "urban_done": ud, "urban_target": ut,
            "rural_done": rd, "rural_target": rt,
            "total_done": td, "total_target": tt,
            "touched": touched,
            "pct": pct, "status": status,
        })
    # sort: started (touched/completed) first (by completion desc), then not started
    rows.sort(key=lambda x: (not x["touched"], -x["total_done"], -x["pct"]))

    mauzas_started = sum(1 for x in rows if x["touched"])
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
    knows_land   = vc(comp["s2_q1"], lab("s2_q1"), dropna=False)
    land_size    = vc(comp["s2_q2"], lab("s2_q2"), dropna=False)

    decision_money = vc(comp["s4_q1"], lab("s4_q1"), dropna=False)
    decision_daily = vc(comp["s4_q2"], lab("s4_q2"), dropna=False)
    decision_save  = vc(comp["s4_q3"], lab("s4_q3"), dropna=False)
    feels_safe     = vc(comp["s4_q4"], lab("s4_q4"), dropna=False)

    # ---- Respondent enrichment ----
    age = pd.to_numeric(comp["s1_q6"], errors="coerce")
    age_marriage = pd.to_numeric(comp["s1_q9"], errors="coerce")
    n_child = pd.to_numeric(comp["no_of_child"], errors="coerce")
    literacy = vc(comp["s1_q14"], lab("s1_q14"), dropna=False)
    earns = vc(comp["s1_q16"], lab("s1_q16"), dropna=False)
    always_village = vc(comp["s1_q7"], lab("s1_q7"), dropna=False)

    # education grouped into readable bands
    edu_raw = pd.to_numeric(comp["s1_q15"], errors="coerce")
    edu_bands = [
        ("No Education", edu_raw == 1),
        ("Primary (1–5)", edu_raw.isin([2, 3, 4, 5, 6])),
        ("Middle (6–8)", edu_raw.isin([7, 8, 9])),
        ("Matric (9–10)", edu_raw.isin([10, 11])),
        ("Intermediate", edu_raw == 12),
        ("Bachelors+", edu_raw.isin([13, 14])),
        ("Other / Madrassa", edu_raw.isin([15, 777])),
    ]
    education = [{"label": l, "value": int(c.sum())} for l, c in edu_bands if int(c.sum()) > 0]

    age_bins = []
    for label, cond in [("18–29", (age >= 18) & (age < 30)), ("30–39", (age >= 30) & (age < 40)),
                        ("40–49", (age >= 40) & (age < 50)), ("50–59", (age >= 50) & (age < 60)),
                        ("60+", age >= 60)]:
        age_bins.append({"label": label, "value": int(cond.sum())})

    avg_age = round(float(age.mean()), 1) if len(age.dropna()) else 0
    avg_age_marriage = round(float(age_marriage.mean()), 1) if len(age_marriage.dropna()) else 0
    avg_children = round(float(n_child.mean()), 1) if len(n_child.dropna()) else 0
    literacy_rate = next((int(round(100 * x["value"] / max(sum(z["value"] for z in literacy if z["label"] != "No response"), 1)))
                          for x in literacy if x["label"] == "Yes"), 0)

    # ---- Vignette (Ayesha vs Amir land-inheritance dispute) ----
    vig_scale = {1: "Very Justified", 2: "Somewhat Justified", 3: "Neutral",
                 4: "Somewhat Unjustified", 5: "Very Unjustified"}
    vignette = {
        "ayesha_sell":  vc(comp["ayesha_sell_justified"], vig_scale, dropna=False),
        "amir_refuse":  vc(comp["amir_refuse_justified"], vig_scale, dropna=False),
        "ayesha_patwari": vc(comp["ayesha_patwari"], vig_scale, dropna=False),
        "ayesha_legal": vc(comp["ayesha_legal"], vig_scale, dropna=False),
        "patwari_side": vc(comp["patwari_side"], lab("patwari_side"), dropna=False),
        "court_side":   vc(comp["court_side"], lab("court_side"), dropna=False),
    }
    def pct_support(arr, keys):
        tot = sum(x["value"] for x in arr if x["label"] != "No response")
        s = sum(x["value"] for x in arr if x["label"] in keys)
        return int(round(100 * s / tot)) if tot else 0
    vignette_meta = {
        "ayesha_sell_support": pct_support(vignette["ayesha_sell"], {"Very Justified", "Somewhat Justified"}),
        "amir_refuse_unjust": pct_support(vignette["amir_refuse"], {"Very Unjustified", "Somewhat Unjustified"}),
        "patwari_ayesha": pct_support(vignette["patwari_side"], {"Ayesha"}),
        "court_ayesha": pct_support(vignette["court_side"], {"Ayesha"}),
    }

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
    # INTERVENTION  (Female Land Survey — Enumerator Script)
    # ------------------------------------------------------------------
    # Only baseline-COMPLETED households are eligible for the intervention,
    # so the eligible pool (the denominator) is the de-duplicated set of
    # completed baseline households — both overall and per mouza. We track
    # how many of those eligible households have so far received the
    # intervention treatment visit.
    intervention = build_intervention(comp, n_complete)

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
            "avg_age": avg_age,
            "avg_age_marriage": avg_age_marriage,
            "avg_children": avg_children,
            "literacy_rate": literacy_rate,
            "vignette": vignette_meta,
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
        "knows_land": knows_land,
        "land_size": land_size,
        "decision_money": decision_money,
        "decision_daily": decision_daily,
        "decision_save": decision_save,
        "feels_safe": feels_safe,
        "mobility": mobility,
        "relationship": rel,
        "education": education,
        "age_bins": age_bins,
        "literacy": literacy,
        "earns": earns,
        "always_village": always_village,
        "vignette": vignette,
        "mouza_table": rows,
        "intervention": intervention,
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
    if intervention["available"]:
        im = intervention["meta"]
        print(f"     Intervention done    : {im['n_iv']} / {im['eligible']} eligible "
              f"({im['pct']}%)  ·  T1 {im['n_t1']} / T2 {im['n_t2']}  ·  "
              f"{im['mauzas_reached']} mauzas reached")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("[ERROR]", e, file=sys.stderr)
        raise
