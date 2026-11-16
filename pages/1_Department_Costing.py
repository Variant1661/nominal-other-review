from __future__ import annotations

import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from department_costing import (  # noqa: E402
    DEFAULT_DEPARTMENTS,
    Rule,
    build_costing,
    load_master_roles,
    normalise_weights,
    read_cost_lines,
    rule_from_dict,
    rules_from_json,
    rules_to_json,
    validate_rules,
    write_costing_workbook,
)
from nominal_logic import DEFAULT_HOME_MAPPING  # noqa: E402

st.set_page_config(page_title="Department Costing", layout="wide")
st.title("Department costing - transfer between homes")
st.caption(
    "Run after wages are posted. Shared departments and shared staff are re-split between homes "
    "by percentage (or days). The output is the net journal Finance needs to post."
)


def parse_mapping(text: str) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for raw in text.splitlines():
        if raw.strip():
            left, _, right = raw.partition("=")
            mapping[left.strip()] = right.strip().upper()
    return mapping


mapping = parse_mapping(st.session_state.get("mapping_text", "")) or dict(DEFAULT_HOME_MAPPING)
homes = [h for _, h in sorted(mapping.items())]

c1, c2 = st.columns(2)
payroll_file = c1.file_uploader("Nominal Update Report (xlsx/csv)", type=["xlsx", "xlsm", "csv"], key="dc_payroll")
master_file = c2.file_uploader("Staff master (xlsx) - needs Full Name, Home, Job Role", type=["xlsx", "xlsm"], key="dc_master")
if not payroll_file:
    st.info("Upload the Nominal Update Report to begin.")
    st.stop()


def save(upload: Any, folder: Path, stem: str) -> Path:
    path = folder / f"{stem}{Path(upload.name).suffix or '.xlsx'}"
    path.write_bytes(upload.getvalue())
    return path


with tempfile.TemporaryDirectory() as tmp:
    tmp_dir = Path(tmp)
    try:
        lines, no_code = read_cost_lines(save(payroll_file, tmp_dir, "payroll"), mapping)
        master, master_df = load_master_roles(save(master_file, tmp_dir, "master") if master_file else None)
    except Exception as exc:  # shown to the user, nothing is silently skipped
        st.error(f"Could not read the files: {exc}")
        st.stop()

if not master_file:
    st.warning("No staff master uploaded: job-role rules and person rules by name cannot be matched.")

# ------------------------------------------------------------- default rules
if "dc_rules" not in st.session_state:
    seed: List[Dict[str, Any]] = [
        {"name": "Kitchen - Crescent/Peter's", "kind": "pool", "suffixes": "05", "job_roles": "", "paying_homes": "CRE,PET", "weights": {"CRE": 50, "PET": 50}},
        {"name": "Admin - Crescent/Peter's", "kind": "pool", "suffixes": "", "job_roles": "Administrator", "paying_homes": "CRE,PET", "weights": {"CRE": 50, "PET": 50}},
    ]
    st.session_state.dc_rules = [rule_from_dict(d) for d in seed]

# Suggest the coordinator rule once, as soon as the staff master is available.
if not master_df.empty and not st.session_state.get("dc_ncc_seeded"):
    st.session_state.dc_ncc_seeded = True
    ncc = master_df[(master_df["Job Role"] == "Non Clinical Coordinator") & (master_df["Home"] == "CRE")]
    if len(ncc) == 1 and not any(r.kind == "person" for r in st.session_state.dc_rules):
        st.session_state.dc_rules.append(rule_from_dict({
            "name": "Non Clinical Coordinator - 2/2/1 days", "kind": "person",
            "person": ncc.iloc[0]["Staff Name"], "weights": {"PET": 2, "CRE": 2, "BUN": 1}}))

rules: List[Rule] = st.session_state.dc_rules

st.subheader("1. Rules")
st.caption("Shares are normalised, so enter percentages (50 / 50) or days (2 / 2 / 1 = 40% / 40% / 20%). Leave a home blank or 0 to exclude it.")

pool_rows = [{"Rule": r.name, "Dept suffixes": ", ".join(r.suffixes), "Job roles": ", ".join(r.job_roles),
              "Paying homes": ", ".join(r.paying_homes), **{h: r.weights.get(h, 0.0) for h in homes}}
             for r in rules if r.kind == "pool"]
person_rows = [{"Rule": r.name, "Staff": r.person, **{h: r.weights.get(h, 0.0) for h in homes}} for r in rules if r.kind == "person"]
share_cfg = {h: st.column_config.NumberColumn(h, min_value=0.0, format="%.2f") for h in homes}

st.markdown("**Shared departments / roles (pool rules)** - all matching cost is pooled, then re-split.")
pool_df = st.data_editor(pd.DataFrame(pool_rows, columns=["Rule", "Dept suffixes", "Job roles", "Paying homes", *homes]),
                         num_rows="dynamic", hide_index=True, width="stretch", column_config=share_cfg, key="dc_pool_editor")

staff_options = sorted(master_df["Staff Name"].tolist()) if not master_df.empty else []
st.markdown("**Individual staff working across homes (person rules)** - overrides any pool rule for that person.")
person_df = st.data_editor(
    pd.DataFrame(person_rows, columns=["Rule", "Staff", *homes]),
    num_rows="dynamic", hide_index=True, width="stretch", key="dc_person_editor",
    column_config={**share_cfg, "Staff": st.column_config.SelectboxColumn("Staff", options=staff_options)} if staff_options else share_cfg,
)


def txt(v: Any) -> str:
    return "" if v is None or (isinstance(v, float) and v != v) else str(v).strip()


def num(v: Any) -> float:
    try:
        f = float(v)
    except (TypeError, ValueError):
        return 0.0
    return f if f == f else 0.0


def rows_to_rules() -> List[Rule]:
    out: List[Rule] = []
    for _, r in pool_df.iterrows():
        if not txt(r.get("Rule")):
            continue
        out.append(rule_from_dict({"name": txt(r["Rule"]), "kind": "pool", "suffixes": txt(r.get("Dept suffixes")),
                                   "job_roles": txt(r.get("Job roles")), "paying_homes": txt(r.get("Paying homes")),
                                   "weights": {h: num(r.get(h)) for h in homes}}))
    for _, r in person_df.iterrows():
        if not txt(r.get("Rule")):
            continue
        out.append(rule_from_dict({"name": txt(r["Rule"]), "kind": "person", "person": txt(r.get("Staff")),
                                   "weights": {h: num(r.get(h)) for h in homes}}))
    return out


current = rows_to_rules()
errors = validate_rules(current, mapping)
if errors:
    st.error("\n".join(errors))
    st.stop()
for r in current:
    st.caption(f"{r.name}: " + ", ".join(f"{h} {s * 100:.1f}%" for h, s in normalise_weights(r.weights).items()))
st.session_state.dc_rules = current

rc1, rc2 = st.columns(2)
rc1.download_button("Save rules (JSON)", rules_to_json(current), "department_costing_rules.json", "application/json")
uploaded_rules = rc2.file_uploader("Load saved rules", type=["json"], key="dc_rules_upload")
if uploaded_rules is not None and st.session_state.get("dc_rules_loaded") != uploaded_rules.file_id:
    try:
        st.session_state.dc_rules = rules_from_json(uploaded_rules.getvalue().decode("utf-8"))
        st.session_state.dc_rules_loaded = uploaded_rules.file_id
        st.rerun()
    except Exception as exc:
        st.error(f"Could not load rules: {exc}")

with st.expander("Department suffix reference (editable meaning is documentation only)"):
    st.write(pd.DataFrame({"Suffix": list(DEFAULT_DEPARTMENTS), "Department": list(DEFAULT_DEPARTMENTS.values())}))
    st.caption("Suffix = nominal code without its first digit, e.g. 205 = Crescent (2) + suffix 05 (Kitchen).")

# ------------------------------------------------------------------ results
st.subheader("2. Result")
result = build_costing(lines, current, mapping, master)
jr, summary = result["journal"], result["summary"]
m1, m2, m3 = st.columns(3)
m1.metric("Payroll lines read", f"{len(lines):,}")
m2.metric("Journal lines", f"{len(jr):,}")
m3.metric("Total moved between homes", f"£{summary['Amount'].sum():,.2f}" if not summary.empty else "£0.00")

if no_code:
    st.warning(f"{len(no_code)} payroll line(s) have no nominal code in the report and were not allocated.")
if not result["warnings"].empty:
    st.warning("Check these before posting:")
    st.dataframe(result["warnings"], hide_index=True, width="stretch")

tab_sum, tab_j, tab_d, tab_p = st.tabs(["Transfer summary", "Journal", "Allocation detail", "Staff included"])
tab_sum.dataframe(summary, hide_index=True, width="stretch")
tab_j.dataframe(jr, hide_index=True, width="stretch")
tab_d.dataframe(result["detail"], hide_index=True, width="stretch")
tab_p.dataframe(result["people"], hide_index=True, width="stretch")

if st.button("Build Excel report", type="primary"):
    with tempfile.TemporaryDirectory() as tmp:
        out = write_costing_workbook(Path(tmp) / "department_costing.xlsx", result, current, no_code, mapping)
        st.session_state.dc_report = out.read_bytes()
if st.session_state.get("dc_report"):
    st.download_button("Download Excel report", st.session_state.dc_report, "department_costing.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
