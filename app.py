"""
Invoice Validator — Fresh clean build for Streamlit Cloud
Upload your Aptive and VACube Excel files through the browser to get started.
"""

import streamlit as st
import pandas as pd
import re
import hashlib
from io import BytesIO
from datetime import datetime
from html import escape
from collections import defaultdict

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Invoice Validator", layout="wide", initial_sidebar_state="expanded")
st.markdown("""<style>
#MainMenu{visibility:hidden;}footer{visibility:hidden;}
.block-container{padding-top:0.5rem !important;}
.stButton>button{border-radius:6px;}
</style>""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
DATE_MISMATCH_REASON = "The Aptive Task Complete Date does not match the service-specific VACube comparison date"
ODS_MISMATCH_REASON = "The Aptive ODS Name does not match the service-specific VACube staff field"
FACILITY_ACCESSION_MISMATCH_REASON = "The Aptive Accession VAMC and Sequence Number do not match the VACube"
VACUBE_LINE_ITEM_DUPLICATE_REASON = "Another Aptive row with the same Accession Sequence and Service is already aligned to the same VACube line item"
DEFAULT_OVERRIDE_REASON = "Validated after manual review of the rejected case details."
APTIVE_INVOICE_SCOPE_START_DATE = "2026-06-16"
PROGRAM_REPORT_UPDATE_MONTHLY_SERVICE = "Program Report Update Monthly"
PROGRAM_REPORT_UPDATE_MONTHLY_CLIN = "005"
PROGRAM_REPORT_UPDATE_MONTHLY_UNIT = 1789.15
SERVICE_UNIT_AMOUNTS = {
    ("casefinding/case ascertainment", "001"): 27.67,
    ("data curation/abstraction", "002"): 98.56,
    ("disease surveillance/follow-up", "003"): 19.14,
    ("program report update monthly", "005"): PROGRAM_REPORT_UPDATE_MONTHLY_UNIT,
}
VACUBE_ODS_STAFF_COLUMNS = ["abstractedby", "caselastchangedby", "initiatedby"]
VACUBE_ODS_EXCLUDED_PAIRS = [
    ("ABDULLA,LAANA A", "(V10) (506) Ann Arbor, MI HCS"),
    ("BAKER,JENNIFER L", "(V22) (691) Greater Los Angeles, CA HCS"),
    ("BRYANT,BEVERLY A", "(V22) (691) Greater Los Angeles, CA HCS"),
    ("CHURCH,DEBRA", "(V22) (691) Greater Los Angeles, CA HCS"),
    ("CHURCH,DEBRA J", "(V15) (589A5) Eastern Kansas HCS"),
    ("CHURCH,DEBRA J", "(V15) (589A7) Wichita, KS HCS"),
    ("CHURCH,DEBRA J", "(V15) (589) Kansas City, MO HCS"),
    ("FIELDS,JULIE A", "(V05) (688) Washington, DC HCS"),
    ("FIELDS,JULIE A", "(V06) (558) Durham, NC HCS"),
    ("FIELDS,JULIE A", "(V06) (658) Salem, VA HCS"),
    ("FIELDS,JULIE A", "(V07) (508) Atlanta, GA HCS"),
    ("FIELDS,JULIE A", "(V17) (671) San Antonio, TX HCS"),
    ("FIELDS,JULIE A", "(V17) (756) El Paso, TX HCS"),
    ("FIELDS,JULIE A", "(V20) (663) Puget Sound, WA HCS"),
    ("FLEMING,CAROL A", "(V04) (642) Philadelphia, PA HCS"),
    ("FLEMING,CAROL A", "(V15) (589) Kansas City, MO HCS"),
    ("FLEMING,CAROL A", "(V15) (589A5) Eastern Kansas HCS"),
    ("FLEMING,CAROL A", "(V15) (589A7) Wichita, KS HCS"),
    ("FLEMING,CAROL A", "(V22) (605) Loma Linda, CA HCS"),
    ("FLEMING,CAROL A", "(V22) (664) San Diego, CA HCS"),
    ("JOHNSON,MALEAH S", "(V01) (689) Connecticut HCS"),
    ("JOHNSON,MALEAH S", "(V02) (526) Bronx, NY HCS"),
    ("JOHNSON,MALEAH S", "(V02) (528) Western New York HCS"),
    ("JOHNSON,MALEAH S", "(V02) (632) Northport, NY HCS"),
    ("JOHNSON,MALEAH S", "(V04) (542) Coatesville, PA HCS"),
    ("JOHNSON,MALEAH S", "(V05) (613) Martinsburg, WV HCS"),
    ("JOHNSON,MALEAH S", "(V06) (558) Durham, NC HCS"),
    ("JOHNSON,MALEAH S", "(V08) (573) Gainesville, FL HCS"),
    ("JOHNSON,MALEAH S", "(V09) (614) Memphis, TN HCS"),
    ("JOHNSON,MALEAH S", "(V17) (671) San Antonio, TX HCS"),
    ("JOHNSON,MALEAH S", "(V19) (436) Montana HCS"),
    ("JOHNSON,MALEAH S", "(V19) (554) Aurora, CO HCS"),
    ("JOHNSON,MALEAH S", "(V20) (663) Puget Sound, WA HCS"),
    ("JOHNSON,MALEAH S", "(V22) (501) New Mexico HCS"),
    ("JOHNSON,MALEAH S", "(V23) (618) Minneapolis, MN HCS"),
    ("LEYNES,ROBYN D", "(V09) (614) Memphis, TN HCS"),
    ("MARTIN,KATHRYN L", "(V01) (689) Connecticut HCS"),
    ("MARTIN,KATHRYN L", "(V23) (618) Minneapolis, MN HCS"),
    ("MULLER,CARLOS", "(V01) (689) Connecticut HCS"),
    ("MULLER,CARLOS", "(V09) (626) Middle Tennessee HCS"),
    ("MULLER,CARLOS", "(V23) (618) Minneapolis, MN HCS"),
    ("OZUMBA,CHIDI C", "(V06) (558) Durham, NC HCS"),
    ("OZUMBA,CHIDI C", "(V06) (658) Salem, VA HCS"),
    ("OZUMBA,CHIDI C", "(V17) (671) San Antonio, TX HCS"),
    ("OZUMBA,CHIDI C", "(V20) (648) Portland, OR HCS"),
    ("OZUMBA,CHIDI C", "(V22) (605) Loma Linda, CA HCS"),
    ("PALMER,NANCY L", "(V04) (642) Philadelphia, PA HCS"),
    ("PALMER,NANCY L", "(V07) (534) Charleston, SC HCS"),
    ("PALMER,NANCY L", "(V08) (548) West Palm Beach, FL HCS"),
    ("PALMER,NANCY L", "(V12) (695) Milwaukee, WI HCS"),
    ("PALMER,NANCY L", "(V22) (605) Loma Linda, CA HCS"),
    ("PALMER,NANCY L", "(V22) (691) Greater Los Angeles, CA HCS"),
    ("ROBINSON,YOLANDA Y", "(V06) (658) Salem, VA HCS"),
    ("ROBINSON,YOLANDA Y", "(V10) (539) Cincinnati, OH HCS"),
    ("ROBINSON,YOLANDA Y", "(V22) (501) New Mexico HCS"),
    ("ROBINSON,YOLANDA Y", "(V22) (605) Loma Linda, CA HCS"),
    ("SORTMAN,MARY MELINDA", "(V07) (544) Columbia, SC HCS"),
    ("STURM,KAILEE J", "(V22) (664) San Diego, CA HCS"),
    ("TAYLOR,ROBERT H", "(V06) (637) Asheville, NC HCS"),
    ("TISDALE,TAMMY", "(V10) (506) Ann Arbor, MI HCS"),
    ("WADDELL,CYNTHIA L", "(V19) (554) Aurora, CO HCS"),
    ("WADDELL,CYNTHIA L", "(V19) (660) Salt Lake City, UT HCS"),
    ("WEISS,BRIDGET C", "(V05) (688) Washington, DC HCS"),
    ("WEISS,BRIDGET C", "(V06) (658) Salem, VA HCS"),
    ("WEISS,BRIDGET C", "(V19) (554) Aurora, CO HCS"),
    ("WILLIAMS,BERNICE", "(V02) (526) Bronx, NY HCS"),
    ("WILLIAMS,BERNICE", "(V02) (528A8) Albany, NY HCS"),
    ("WILLIAMS,BERNICE", "(V02) (620) Hudson Valley, NY HCS"),
    ("WILLIAMS,BERNICE", "(V08) (548) West Palm Beach, FL HCS"),
    ("WILLIAMS,BERNICE", "(V10) (553) Detroit, MI HCS"),
    ("WILLIAMS,BERNICE", "(V22) (691) Greater Los Angeles, CA HCS"),
]

# ── Normalization helpers ─────────────────────────────────────────────────────

def norm_facility(v):
    if pd.isna(v): return ""
    s = str(v).strip().lower()
    if s in ("", "nan", "none", "not entered"): return ""
    s = re.sub(r"^\(v\d+\)\s*", "", s)
    return re.sub(r"\s+", " ", s).strip()

def norm_person(v):
    if pd.isna(v): return ""
    s = str(v).strip().lower()
    if s in ("", "nan", "none", "not entered"): return ""
    s = s.replace(".", " ")
    return " ".join(s.replace(",", " , ").split())

def person_keys(v):
    n = norm_person(v)
    if not n: return set()
    keys = {n.replace(" ", "")}
    if "," in n:
        last, first = [p.strip() for p in n.split(",", 1)]
        lt = [t for t in last.split() if t.isalpha()]
        ft = [t for t in first.split() if t.isalpha()]
        if lt and ft: keys.add(f"{ft[0]}|{' '.join(lt)}")
    else:
        toks = [t for t in n.split() if t.isalpha()]
        if len(toks) >= 2: keys.add(f"{toks[0]}|{toks[-1]}")
    return {k for k in keys if k}

def names_match(a, b):
    return bool(person_keys(a) & person_keys(b))

def fmt_clin(v):
    if pd.isna(v): return "Unknown"
    s = str(v).strip()
    if s in ("", "nan", "None"): return "Unknown"
    try:
        n = float(s)
        if n.is_integer():
            i = int(n)
            return f"{i:03d}" if 0 <= i <= 999 else str(i)
    except (TypeError, ValueError): pass
    return s

def norm_service(v):
    if pd.isna(v): return "unknown"
    return str(v).strip().lower() or "unknown"

def accession9(v):
    return re.sub(r"\s+", "", str(v) if not pd.isna(v) else "")[:9]

# ── Core matching engine ──────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def run_comparison(_aptive_bytes, _vacube_bytes, date_from_str):
    """Full matching pipeline. Returns a results dict."""
    aptive_df = pd.read_excel(BytesIO(_aptive_bytes))
    vacube_df = pd.read_excel(BytesIO(_vacube_bytes))

    # Normalize columns
    aptive_df.columns = aptive_df.columns.str.strip().str.lower()
    vacube_df.columns = vacube_df.columns.str.strip().str.lower()
    aptive_df = aptive_df.reset_index(drop=True)
    vacube_df = vacube_df.reset_index(drop=True)

    # Check required columns
    missing_a = [c for c in ["vamc", "accession sequence #", "task complete date"] if c not in aptive_df.columns]
    missing_v = [c for c in ["accession_number"] if c not in vacube_df.columns]
    if missing_a or missing_v:
        return {"error": f"Missing Aptive columns: {missing_a}. Missing VACube columns: {missing_v}.",
                "aptive_cols": list(aptive_df.columns), "vacube_cols": list(vacube_df.columns)}

    # Date scope filter
    excluded_count = 0
    date_from = pd.to_datetime(date_from_str, errors="coerce").normalize() if date_from_str else None
    if date_from and "date invoiced" in aptive_df.columns:
        inv_dates = pd.to_datetime(aptive_df["date invoiced"], errors="coerce").dt.normalize()
        mask = inv_dates.ge(date_from)
        excluded_count = int((~mask).sum())
        aptive_df = aptive_df[mask].reset_index(drop=True)
        date_counts = inv_dates[mask].dropna().dt.strftime("%Y-%m-%d").value_counts().sort_index().to_dict()
        latest_date = inv_dates[mask].dropna().max()
        latest_date_str = latest_date.strftime("%Y-%m-%d") if pd.notna(latest_date) else ""
    else:
        date_counts = {}
        latest_date_str = ""

    # Pre-compute fields
    aptive_df["_acc9"] = aptive_df["accession sequence #"].apply(accession9)
    aptive_df["_vamc_n"] = aptive_df["vamc"].apply(norm_facility)
    aptive_df["_svc_n"] = aptive_df.get("service", pd.Series("unknown", index=aptive_df.index)).apply(norm_service)
    aptive_df["_clin_n"] = aptive_df.get("task item / clin", pd.Series("Unknown", index=aptive_df.index)).apply(fmt_clin)
    aptive_df["_taskdate"] = pd.to_datetime(aptive_df["task complete date"], errors="coerce").dt.date
    aptive_df["_ods"] = aptive_df.get("ods name", pd.Series("", index=aptive_df.index)).fillna("")

    vacube_df["_acc9"] = vacube_df["accession_number"].apply(accession9)
    vacube_df["_fac_n"] = vacube_df.get("facility1", pd.Series("", index=vacube_df.index)).apply(norm_facility)
    vacube_df["_date_init"] = pd.to_datetime(vacube_df.get("datecaseinitiated", pd.NaT), errors="coerce").dt.date
    vacube_df["_date_chg"] = pd.to_datetime(vacube_df.get("datecaselastchanged1", pd.NaT), errors="coerce").dt.date
    vacube_df["_date_cmp"] = pd.to_datetime(vacube_df.get("datecasecompleted1", pd.NaT), errors="coerce").dt.date
    vacube_df["_abs_status"] = vacube_df.get("abstractstatus", pd.Series("", index=vacube_df.index)).fillna("").str.lower()

    # Build VACube index by accession9
    vac_index = defaultdict(list)
    for i, row in vacube_df.iterrows():
        vac_index[row["_acc9"]].append(row)

    matched_rows = []
    rejected_rows = []
    used_vac_keys = set()  # (vac_row_id, acc9_norm, svc_key) to prevent duplicate VACube line items

    for idx, arow in aptive_df.iterrows():
        acc9 = arow["_acc9"]
        vamc_n = arow["_vamc_n"]
        svc_n = arow["_svc_n"]
        clin_n = arow["_clin_n"]
        task_date = arow["_taskdate"]
        ods = arow["_ods"]

        candidates = vac_index.get(acc9, [])

        # Determine service type
        is_cf = svc_n == "casefinding/case ascertainment" or clin_n == "001"
        is_ab = svc_n == "data curation/abstraction" or clin_n == "002"
        is_sv = svc_n == "disease surveillance/follow-up" or clin_n == "003"

        if not candidates:
            rejected_rows.append({**arow.to_dict(), "Reason": FACILITY_ACCESSION_MISMATCH_REASON})
            continue

        # Try to find a match
        best = None
        fac_matched = False
        date_matched = False
        ods_matched = False

        for vrow in candidates:
            fac_ok = vrow["_fac_n"] == vamc_n

            # Service-specific date
            if is_cf: vdate = vrow["_date_init"]
            elif is_ab: vdate = vrow["_date_cmp"]
            elif is_sv: vdate = vrow["_date_chg"]
            else: vdate = vrow["_date_init"] or vrow["_date_chg"] or vrow["_date_cmp"]

            date_ok = task_date == vdate if (task_date and vdate) else False

            # ODS/staff match
            ods_str = str(ods).strip()
            if not ods_str:
                ods_ok = True
            else:
                if is_cf: staff = vrow.get("initiatedby", "")
                elif is_ab: staff = vrow.get("abstractedby", "")
                elif is_sv: staff = vrow.get("caselastchangedby", "")
                else: staff = vrow.get("initiatedby", "") or vrow.get("abstractedby", "") or vrow.get("caselastchangedby", "")
                ods_ok = names_match(ods_str, str(staff))

            # CLIN 002 requires abstract status complete
            if is_ab and vrow["_abs_status"] != "complete":
                continue

            if date_ok and ods_ok and (fac_ok or (date_ok and ods_ok)):
                svc_key = svc_n if svc_n not in ("", "unknown") else clin_n
                vac_key = (id(vrow), acc9, svc_key)
                if vac_key in used_vac_keys:
                    rejected_rows.append({**arow.to_dict(), "Reason": VACUBE_LINE_ITEM_DUPLICATE_REASON})
                    best = None
                    break
                best = vrow
                fac_matched = fac_ok
                date_matched = date_ok
                ods_matched = ods_ok
                used_vac_keys.add(vac_key)
                break

        if best is not None:
            merged = {**arow.to_dict()}
            for k, v in best.items():
                if not k.startswith("_"): merged[f"vac_{k}"] = v
            merged["facility_matched"] = fac_matched
            matched_rows.append(merged)
        elif not any(r.get("aptive_row_id") == idx for r in rejected_rows):
            # Determine specific rejection reason
            has_acc = bool(candidates)
            fac_acc_ok = any(c["_fac_n"] == vamc_n for c in candidates)
            if is_cf: dates_ok = any(task_date == c["_date_init"] for c in candidates if task_date)
            elif is_ab: dates_ok = any(task_date == c["_date_cmp"] for c in candidates if task_date)
            elif is_sv: dates_ok = any(task_date == c["_date_chg"] for c in candidates if task_date)
            else: dates_ok = False

            if fac_acc_ok and dates_ok:
                reason = ODS_MISMATCH_REASON
            elif fac_acc_ok:
                reason = DATE_MISMATCH_REASON
            else:
                reason = FACILITY_ACCESSION_MISMATCH_REASON
            rejected_rows.append({**arow.to_dict(), "Reason": reason})

    matched_df = pd.DataFrame(matched_rows)
    rejected_df = pd.DataFrame(rejected_rows)

    # Rejection breakdown
    rej_breakdown = {}
    reason_code = {FACILITY_ACCESSION_MISMATCH_REASON: "R1", DATE_MISMATCH_REASON: "R2", ODS_MISMATCH_REASON: "R3", VACUBE_LINE_ITEM_DUPLICATE_REASON: "R1"}
    if not rejected_df.empty and "Reason" in rejected_df.columns:
        for r in rejected_df["Reason"].fillna("").astype(str):
            code = reason_code.get(r, "R1")
            rej_breakdown[code] = rej_breakdown.get(code, 0) + 1

    # Facility breakdown
    fac_breakdown = []
    if not matched_df.empty or not rejected_df.empty:
        all_df = pd.concat([
            matched_df.assign(_status="matched") if not matched_df.empty else pd.DataFrame(),
            rejected_df.assign(_status="rejected") if not rejected_df.empty else pd.DataFrame(),
        ], ignore_index=True, sort=False)
        if "vamc" in all_df.columns:
            for vamc, grp in all_df.groupby("vamc"):
                total = len(grp); mat = int((grp["_status"] == "matched").sum())
                fac_breakdown.append({"vamc": vamc, "total": total, "matched": mat, "rejected": total - mat, "rate": round(mat / total * 100, 1) if total else 0})
            fac_breakdown.sort(key=lambda x: x["rate"])

    # Validated services summary
    svc_summary = []
    if not matched_df.empty and "service" in matched_df.columns:
        matched_df["_clin_fmt"] = matched_df.get("task item / clin", pd.Series("Unknown", index=matched_df.index)).apply(fmt_clin)
        matched_df["_svc_fmt"] = matched_df["service"].fillna("Unknown").astype(str).str.strip()
        matched_df["_amt"] = pd.to_numeric(matched_df.get("amt. invoiced", pd.Series(0, index=matched_df.index)).apply(str).str.replace(r"[$,]", "", regex=True), errors="coerce").fillna(0)
        for (svc, clin), grp in matched_df.groupby(["_svc_fmt", "_clin_fmt"]):
            unit = SERVICE_UNIT_AMOUNTS.get((norm_service(svc), clin), None)
            cases = len(grp)
            total_amt = cases * unit if unit else grp["_amt"].sum()
            svc_summary.append({"Service": svc, "Task Item / CLIN": clin, "Cases": cases, "Total Amt. Invoiced": f"${total_amt:,.2f}", "_sort": total_amt})
        svc_summary.sort(key=lambda x: -x["_sort"])
        for s in svc_summary: del s["_sort"]

    total_in_scope = len(aptive_df)
    total_matched = len(matched_df)
    match_rate = round(total_matched / total_in_scope * 100, 1) if total_in_scope else 0.0
    validated_amt = sum(float(s["Total Amt. Invoiced"].replace("$", "").replace(",", "")) for s in svc_summary)

    return {
        "matched_df": matched_df,
        "rejected_df": rejected_df,
        "summary": {
            "aptive_total": total_in_scope,
            "vacube_total": len(vacube_df),
            "matched": total_matched,
            "rejected": len(rejected_df),
            "match_rate": match_rate,
            "validated_amt": validated_amt,
            "excluded": excluded_count,
            "date_from": date_from_str,
            "latest_date": latest_date_str,
            "date_counts": date_counts,
        },
        "rej_breakdown": rej_breakdown,
        "fac_breakdown": fac_breakdown,
        "svc_summary": svc_summary,
    }

# ── UI helpers ────────────────────────────────────────────────────────────────

def render_topbar(run_id=""):
    st.markdown(f"""
    <div style="background:#fff;border-bottom:1px solid #e5e7eb;padding:10px 0 8px 0;margin-bottom:1.2em;display:flex;align-items:center;justify-content:space-between;">
      <div style="display:flex;align-items:center;gap:12px;">
        <div style="width:34px;height:34px;background:#185FA5;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:13px;font-weight:700;">IV</div>
        <div>
          <div style="font-size:16px;font-weight:700;color:#111;">Invoice Validator</div>
          <div style="font-size:10px;color:#888;letter-spacing:.4px;">VA CANCER REGISTRY · APTIVE → VACUBE</div>
        </div>
      </div>
      {"<span style='font-size:11px;color:#185FA5;background:#eff6ff;padding:3px 10px;border-radius:20px;font-family:monospace;'>run · #" + run_id + "</span>" if run_id else ""}
    </div>""", unsafe_allow_html=True)

def render_metric_cards(cards):
    cols = st.columns(len(cards))
    for col, card in zip(cols, cards):
        label, value = card[0], card[1]
        color = card[2] if len(card) > 2 else "#185FA5"
        with col:
            st.markdown(f"""
            <div style="background:#fff;border:1px solid #e5e7eb;border-radius:8px;padding:14px 10px;text-align:center;min-height:88px;display:flex;flex-direction:column;justify-content:center;">
              <div style="font-size:10px;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px;">{escape(str(label))}</div>
              <div style="font-size:22px;font-weight:700;color:{color};">{escape(str(value))}</div>
            </div>""", unsafe_allow_html=True)

def render_alert(severity, headline, detail=""):
    cm = {"high": ("#FAECE7","#D85A30","#993C1D","HIGH"), "medium": ("#FAEEDA","#BA7517","#854F0B","MEDIUM"), "info": ("#eff6ff","#93c5fd","#185FA5","INFO"), "success": ("#f0fdf4","#86efac","#166534","OK")}
    bg, border, text, sev_label = cm.get(severity, cm["info"])
    st.markdown(f"""
    <div style="background:{bg};border-left:3px solid {border};padding:10px 14px;border-radius:0 8px 8px 0;margin-bottom:8px;font-size:13px;">
      <span style="font-size:10px;font-weight:700;letter-spacing:.5px;color:{border};text-transform:uppercase;margin-right:6px;">{sev_label}</span>
      <strong style="color:{text};">{escape(str(headline))}</strong>
      {f'<div style="color:#6b7280;margin-top:3px;">{escape(str(detail))}</div>' if detail else ''}
    </div>""", unsafe_allow_html=True)

def render_rejection_chart(rej_breakdown):
    LABELS = {"R1": "Accession mismatch", "R2": "Date mismatch", "R3": "Staff mismatch"}
    COLORS = {"R1": "#E24B4A", "R2": "#D85A30", "R3": "#7F77DD"}
    if not rej_breakdown: st.caption("No rejection data."); return
    total = sum(rej_breakdown.values()) or 1
    bars = ""
    for code, cnt in sorted(rej_breakdown.items(), key=lambda x: -x[1]):
        pct = int(cnt / total * 100)
        bars += f"""<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
          <div style="font-size:11px;color:#6b7280;width:130px;flex-shrink:0;text-align:right;">{escape(LABELS.get(code, code))}</div>
          <div style="flex:1;height:16px;background:#f3f4f6;border-radius:4px;overflow:hidden;"><div style="width:{pct}%;height:100%;background:{COLORS.get(code,'#888')};border-radius:4px;"></div></div>
          <div style="font-size:11px;color:#6b7280;width:28px;text-align:right;">{cnt}</div>
        </div>"""
    st.markdown(f"""<div style="background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;">
      <div style="font-size:10px;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:12px;">Rejection reason breakdown</div>{bars}</div>""", unsafe_allow_html=True)

def render_facility_table(fac_breakdown):
    if not fac_breakdown: return
    def pill(r):
        if r >= 75: return f"<span style='background:#EAF3DE;color:#3B6D11;padding:2px 7px;border-radius:20px;font-size:11px;font-weight:500;'>{r}%</span>"
        if r >= 55: return f"<span style='background:#FAEEDA;color:#854F0B;padding:2px 7px;border-radius:20px;font-size:11px;font-weight:500;'>{r}%</span>"
        return f"<span style='background:#FAECE7;color:#993C1D;padding:2px 7px;border-radius:20px;font-size:11px;font-weight:500;'>{r}%</span>"
    rows = "".join(f"<tr><td style='padding:7px 8px;color:#111;font-weight:500;font-size:12px;'>{escape(str(r['vamc']))}</td><td style='padding:7px 8px;text-align:right;color:#6b7280;font-size:12px;'>{r['total']}</td><td style='padding:7px 8px;text-align:right;color:#6b7280;font-size:12px;'>{r['matched']}</td><td style='padding:7px 8px;text-align:right;color:#6b7280;font-size:12px;'>{r['rejected']}</td><td style='padding:7px 8px;text-align:right;'>{pill(r['rate'])}</td></tr>" for r in fac_breakdown[:8])
    st.markdown(f"""<div style="background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;">
      <div style="font-size:10px;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:12px;">Facility-level match rate</div>
      <table style="width:100%;border-collapse:collapse;">
        <thead><tr>{"".join(f'<th style="font-size:10px;color:#9ca3af;text-transform:uppercase;padding:4px 8px;font-weight:400;text-align:{"left" if h=="VAMC" else "right"}">{h}</th>' for h in ["VAMC","Total","Matched","Rejected","Rate"])}</tr></thead>
        <tbody>{rows}</tbody>
      </table></div>""", unsafe_allow_html=True)

def render_ai_insights(rej_breakdown, fac_breakdown, summary):
    insights = []
    total_rej = summary.get("rejected", 0) or 1
    if rej_breakdown:
        top_code, top_cnt = max(rej_breakdown.items(), key=lambda x: x[1])
        pct = round(top_cnt / total_rej * 100)
        tips = {"R2": "Consider checking for timezone or batch-processing lag in the Aptive submission.", "R1": "Verify accession number formatting between Aptive and VACube.", "R3": "Review ODS name spelling consistency between systems."}
        lbls = {"R2": "date mismatches", "R1": "accession mismatches", "R3": "staff mismatches"}
        insights.append(f"{pct}% of rejections ({top_cnt} cases) are {lbls.get(top_code, top_code)} ({top_code}). {tips.get(top_code, '')}")
    if fac_breakdown:
        worst = fac_breakdown[0]
        insights.append(f"VAMC {worst['vamc']} has the lowest match rate at {worst['rate']}% ({worst['rejected']} rejected of {worst['total']} total). Prioritize review here.")
    if summary.get("match_rate", 100) < 60:
        insights.append(f"Match rate of {summary.get('match_rate', 0):.1f}% is critically low. Do not submit this invoice batch without resolving key rejections.")
    if not insights: insights.append("No significant issues detected in the current run.")
    items = "".join(f"<li style='font-size:12px;color:#1e3a5f;line-height:1.6;margin-bottom:4px;'>{escape(i)}</li>" for i in insights)
    st.markdown(f"""<div style="background:#f0f7ff;border:1px solid #bfdbfe;border-radius:10px;padding:14px 18px;margin-bottom:16px;">
      <div style="font-size:12px;font-weight:600;color:#185FA5;margin-bottom:10px;">✦ AI insight summary</div>
      <ul style="margin:0;padding-left:16px;">{items}</ul></div>""", unsafe_allow_html=True)

def to_excel_bytes(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

# ── Upload screen ─────────────────────────────────────────────────────────────

def show_upload_screen():
    st.markdown("""
    <div style="max-width:560px;margin:60px auto 0;">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:24px;">
        <div style="width:40px;height:40px;background:#185FA5;border-radius:10px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:15px;font-weight:700;">IV</div>
        <div>
          <div style="font-size:22px;font-weight:700;color:#111;">Invoice Validator</div>
          <div style="font-size:11px;color:#888;">VA Cancer Registry · Aptive → VACube</div>
        </div>
      </div>
      <p style="color:#555;font-size:14px;line-height:1.6;border-top:1px solid #f0f0f0;padding-top:20px;">
        Compare Aptive invoice records against VACube clinical data to validate contractor billing.
        Upload both Excel files below to get started.
      </p>
    </div>
    """, unsafe_allow_html=True)

    _, center, _ = st.columns([1, 2, 1])
    with center:
        aptive_file = st.file_uploader("Aptive Invoice File (.xlsx)", type=["xlsx"], key="aptive_upload")
        vacube_file = st.file_uploader("VACube Export File (.xlsx)", type=["xlsx"], key="vacube_upload")
        date_from = st.date_input(
            "Date Invoiced From (optional — leave blank for all dates)",
            value=datetime.strptime(APTIVE_INVOICE_SCOPE_START_DATE, "%Y-%m-%d").date(),
            key="date_from_input",
        )
        if st.button("Upload & Compare →", use_container_width=True, type="primary"):
            if not aptive_file or not vacube_file:
                st.error("Please upload both files before continuing.")
            else:
                with st.spinner("Running comparison..."):
                    aptive_bytes = aptive_file.read()
                    vacube_bytes = vacube_file.read()
                    result = run_comparison(aptive_bytes, vacube_bytes, date_from.isoformat())
                    if "error" in result:
                        st.error(result["error"])
                        st.code(f"Aptive columns: {result.get('aptive_cols', [])}\nVACube columns: {result.get('vacube_cols', [])}")
                    else:
                        import random
                        st.session_state["result"] = result
                        st.session_state["aptive_bytes"] = aptive_bytes
                        st.session_state["vacube_bytes"] = vacube_bytes
                        st.session_state["run_id"] = "".join(random.choices("0123456789abcdef", k=7))
                        st.session_state["overrides"] = {}
                        st.rerun()

# ── Dashboard ─────────────────────────────────────────────────────────────────

def show_dashboard():
    result = st.session_state["result"]
    summary = result["summary"]
    rej_breakdown = result["rej_breakdown"]
    fac_breakdown = result["fac_breakdown"]
    svc_summary = result["svc_summary"]
    matched_df = result["matched_df"]
    rejected_df = result["rejected_df"]
    overrides = st.session_state.get("overrides", {})
    run_id = st.session_state.get("run_id", "")

    # Apply overrides to rejected df
    if overrides and not rejected_df.empty and "Reason" in rejected_df.columns:
        override_mask = rejected_df.index.isin(overrides.keys())
        override_df = rejected_df[override_mask].copy()
        override_df["Manual Validation"] = "Manual override recorded"
        override_df["Override Reason"] = override_df.index.map(lambda i: overrides.get(i, DEFAULT_OVERRIDE_REASON))
        active_rejected_df = rejected_df[~override_mask].copy()
        combined_matched_df = pd.concat([matched_df, override_df], ignore_index=True, sort=False)
    else:
        active_rejected_df = rejected_df.copy()
        combined_matched_df = matched_df.copy()

    n_matched = len(combined_matched_df)
    n_rejected = len(active_rejected_df)
    n_total = summary["aptive_total"]
    match_rate = round(n_matched / n_total * 100, 1) if n_total else 0.0

    # Recalculate validated amount with overrides
    validated_amt = 0.0
    if not combined_matched_df.empty:
        amt_col = next((c for c in combined_matched_df.columns if "amt" in c.lower() and "invoic" in c.lower()), None)
        if amt_col:
            validated_amt = pd.to_numeric(combined_matched_df[amt_col].apply(str).str.replace(r"[$,]", "", regex=True), errors="coerce").fillna(0).sum()

    with st.sidebar:
        st.markdown("### Navigation")
        page = st.radio("", ["Key Issues & Changes", "Records that Match", "Rejected Cases"], label_visibility="collapsed")
        st.markdown("---")
        st.markdown(f"<div style='font-size:11px;color:#888;'>Date scope</div><div style='font-size:13px;color:#185FA5;font-weight:500;'>{summary.get('date_from','All dates') or 'All dates'}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:11px;color:#888;margin-top:10px;'>Matched</div><div style='font-size:13px;color:#1D9E75;font-weight:700;'>{n_matched:,}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:11px;color:#888;margin-top:6px;'>Rejected</div><div style='font-size:13px;color:#D85A30;font-weight:700;'>{n_rejected:,}</div>", unsafe_allow_html=True)
        st.markdown("---")
        if st.button("↺ New Review", use_container_width=True):
            for k in ["result", "aptive_bytes", "vacube_bytes", "run_id", "overrides"]:
                st.session_state.pop(k, None)
            st.cache_data.clear()
            st.rerun()

    render_topbar(run_id)

    # ── KEY ISSUES ──────────────────────────────────────────────────────────
    if page == "Key Issues & Changes":
        st.markdown("<h2 style='font-size:1.5em;margin-bottom:0.3em;color:#111;'>Key Issues Dashboard</h2>", unsafe_allow_html=True)

        scope_label = f"Date scope from {summary['date_from']}" if summary.get("date_from") else "Full date scope"
        st.caption(f"{scope_label} · {summary['aptive_total']:,} Aptive line items · {summary['vacube_total']:,} VACube records")

        # 6 metric cards
        render_metric_cards([
            ("Match Rate",      f"{match_rate}%",                                  "#D85A30" if match_rate < 70 else "#1D9E75"),
            ("Matches",         f"{n_matched:,}",                                  "#1D9E75"),
            ("Amt. Invoiced",   f"${validated_amt:,.2f}",                          "#374151"),
            ("Duplicates",      "0",                                               "#1D9E75"),
            ("Unmatched",       f"{n_rejected:,} ({100-match_rate:.1f}%)",         "#D85A30"),
            ("VACube Records",  f"{summary['vacube_total']:,}",                    "#378ADD"),
        ])

        # AI insights
        render_ai_insights(rej_breakdown, fac_breakdown, {**summary, "rejected": n_rejected, "match_rate": match_rate})

        # Alerts
        if match_rate < 70:
            top_r = max(rej_breakdown.items(), key=lambda x: x[1]) if rej_breakdown else None
            detail = f"Top rejection driver: {top_r[0]} accounts for {top_r[1]} cases ({round(top_r[1]/n_rejected*100) if n_rejected else 0}% of rejections)." if top_r else ""
            render_alert("high", f"Critical: match rate at {match_rate}% — investigate before invoice approval", detail)
        if summary.get("excluded", 0):
            render_alert("info", f"{summary['excluded']:,} older Aptive rows excluded (before {summary.get('date_from','')})")

        # Charts
        c1, c2 = st.columns(2)
        with c1: render_rejection_chart(rej_breakdown)
        with c2: render_facility_table(fac_breakdown)

        # Validated services
        st.markdown("<h3 style='font-size:1.1em;margin-top:1em;margin-bottom:0.5em;color:#185FA5;'>Validated Services</h3>", unsafe_allow_html=True)
        if overrides:
            st.success(f"Validated Services includes {len(overrides)} case(s) manually overridden from Rejected Cases.")

        if svc_summary:
            svc_df = pd.DataFrame(svc_summary)
            st.dataframe(svc_df[["Service", "Task Item / CLIN", "Cases", "Total Amt. Invoiced"]], use_container_width=True, hide_index=True)
        else:
            st.caption("No matched records with service values found.")

        # Downloads
        st.markdown("---")
        dl1, dl2 = st.columns(2)
        with dl1:
            if not combined_matched_df.empty:
                clean = combined_matched_df[[c for c in combined_matched_df.columns if not c.startswith("_")]].copy()
                st.download_button("📥 Download Matched Records", to_excel_bytes(clean), "matched_records.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with dl2:
            if not active_rejected_df.empty:
                clean = active_rejected_df[[c for c in active_rejected_df.columns if not c.startswith("_")]].copy()
                st.download_button("📥 Download Rejected Cases", to_excel_bytes(clean), "rejected_cases.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        # Overall analysis
        with st.expander("Overall Analysis & Match Rules", expanded=False):
            st.markdown(f"""
- Comparison scope: **{summary['aptive_total']:,}** Aptive records, **{summary['vacube_total']:,}** VACube records.
- **{n_matched:,}** records matched (including **{len(overrides):,}** manual overrides). Match rate: **{match_rate:.1f}%**.
- **{n_rejected:,}** records remain unmatched.
- The largest unmatched driver: **{max(rej_breakdown, key=rej_breakdown.get, default='N/A')}** with **{max(rej_breakdown.values(), default=0)}** cases.

**How a Case Is Determined to Match**
- The Aptive VAMC must match the VACube Facility after normalizing for spaces and capitalization.
- The first 9 digits of the Aptive Accession Sequence Number must match the first 9 digits of the VACube Accession Number.
- The Aptive Task Complete Date must match the VACube date for the service: Casefinding uses Date Case Initiated, Data Curation uses Date Case Completed, Disease Surveillance uses Date Case Last Changed.
- When Aptive ODS Name is present, it must match the corresponding VACube staff field.
- If the facility text does not match but accession, date, and staff all match, the case can still be validated.
            """)

    # ── RECORDS THAT MATCH ──────────────────────────────────────────────────
    elif page == "Records that Match":
        st.markdown("<h2 style='font-size:1.3em;margin-bottom:0.5em;'>Records that Match</h2>", unsafe_allow_html=True)
        render_metric_cards([("Total Matched", f"{n_matched:,}", "#1D9E75")])

        if not combined_matched_df.empty:
            display_cols = [c for c in combined_matched_df.columns if not c.startswith("_")]
            priority = ["vamc", "accession sequence #", "task item / clin", "service", "task complete date", "ods name", "date invoiced", "amt. invoiced"]
            ordered = [c for c in priority if c in display_cols] + [c for c in display_cols if c not in priority]
            search = st.text_input("🔍 Search by VAMC or Accession #", "", key="match_search")
            show_df = combined_matched_df[ordered].copy()
            if search:
                mask = show_df.get("vamc", pd.Series(dtype=str)).astype(str).str.contains(search, case=False, na=False) | show_df.get("accession sequence #", pd.Series(dtype=str)).astype(str).str.contains(search, case=False, na=False)
                show_df = show_df[mask]
            st.dataframe(show_df, use_container_width=True, hide_index=True)
            st.download_button("📥 Download Matched Records", to_excel_bytes(show_df), "matched_records.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("No matched records yet.")

    # ── REJECTED CASES ──────────────────────────────────────────────────────
    elif page == "Rejected Cases":
        st.markdown("<h2 style='font-size:1.3em;margin-bottom:0.5em;'>Rejected Cases</h2>", unsafe_allow_html=True)

        amt_rej = 0.0
        if not active_rejected_df.empty and "amt. invoiced" in active_rejected_df.columns:
            amt_rej = pd.to_numeric(active_rejected_df["amt. invoiced"].apply(str).str.replace(r"[$,]", "", regex=True), errors="coerce").fillna(0).sum()

        render_metric_cards([
            ("Total Rejected", f"{n_rejected:,}", "#D85A30"),
            ("Amt. Invoiced (Rejected)", f"${amt_rej:,.2f}", "#92400e"),
            ("Manual Overrides Applied", f"{len(overrides):,}", "#1D9E75"),
        ])

        if not active_rejected_df.empty:
            display_cols = [c for c in active_rejected_df.columns if not c.startswith("_")]
            priority = ["vamc", "accession sequence #", "task item / clin", "service", "task complete date", "ods name", "Reason"]
            ordered = [c for c in priority if c in display_cols] + [c for c in display_cols if c not in priority]
            reason_filter = st.selectbox("Filter by rejection reason", ["All", "R1 — Accession mismatch", "R2 — Date mismatch", "R3 — Staff mismatch"], key="rej_filter")
            show_df = active_rejected_df[ordered].copy()
            if reason_filter != "All" and "Reason" in show_df.columns:
                code = reason_filter.split(" ")[0]
                code_map = {"R1": FACILITY_ACCESSION_MISMATCH_REASON, "R2": DATE_MISMATCH_REASON, "R3": ODS_MISMATCH_REASON}
                show_df = show_df[show_df["Reason"] == code_map.get(code, "")]
            st.dataframe(show_df, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("**Manual Override** — select cases to move to Records that Match")

            override_reason = st.text_area("Override reason", value=DEFAULT_OVERRIDE_REASON, key="override_reason_input")

            oc1, oc2 = st.columns(2)
            with oc1:
                if st.button("Override All Rejected Cases", use_container_width=True):
                    for i in active_rejected_df.index:
                        st.session_state["overrides"][i] = override_reason
                    st.success(f"All {n_rejected} rejected cases moved to Records that Match.")
                    st.rerun()

            st.markdown("**Or select individual cases:**")
            for i, row in active_rejected_df.iterrows():
                label = f"{row.get('vamc','')} | {row.get('accession sequence #','')} | {row.get('service','')} | {row.get('Reason','')}"
                if st.checkbox(label[:120], key=f"cb_{i}"):
                    if st.button(f"Override case {i}", key=f"ovr_{i}"):
                        st.session_state["overrides"][i] = override_reason
                        st.rerun()

            st.download_button("📥 Download Rejected Cases", to_excel_bytes(show_df), "rejected_cases.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("🎉 No rejected cases — all records have been matched or manually validated.")

# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    if "result" not in st.session_state:
        show_upload_screen()
    else:
        show_dashboard()

if __name__ == "__main__":
    main()
