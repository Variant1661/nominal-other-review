from __future__ import annotations

import hashlib
import tempfile
from pathlib import Path
from typing import Dict, List

import pandas as pd
import streamlit as st

from nominal_logic import (
    DEFAULT_HOME_MAPPING,
    ParsedLine,
    apply_missing_nominals,
    build_review_rows,
    clean_nominal,
    default_output_name,
    infer_staff_base,
    load_staff_master,
    missing_nominal_key,
    nominal_dropdown_options,
    nominal_home,
    read_lines,
    review_lines,
    staff_key,
    validate_nominal,
    write_output,
)


st.set_page_config(page_title="Nominal OTHER Review", page_icon=":bar_chart:", layout="wide")

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.6rem; max-width: 1280px;}
    div[data-testid="stMetricValue"] {font-size: 1.4rem;}
    .stButton button {border-radius: 0.45rem; font-weight: 600;}
    .stDownloadButton button {border-radius: 0.45rem; font-weight: 700;}
    </style>
    """,
    unsafe_allow_html=True,
)


def uploaded_signature(uploaded_file: object) -> str:
    data = uploaded_file.getvalue()
    return hashlib.sha256(data).hexdigest()


def save_uploaded_file(uploaded_file: object, folder: Path, stem: str) -> Path:
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    path = folder / f"{stem}{suffix}"
    path.write_bytes(uploaded_file.getvalue())
    return path


def parse_mapping(text: str) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if "=" not in line:
            raise ValueError(f"Invalid mapping line: {line}. Use format 2=CRE.")
        left, right = [part.strip() for part in line.split("=", 1)]
        if not left.isdigit() or len(left) != 1:
            raise ValueError(f"Invalid nominal prefix: {left}. Use one digit, for example 5=BUN.")
        if not right.isalpha() or not (2 <= len(right) <= 5):
            raise ValueError(f"Invalid home code: {right}. Use short codes like CRE, PET, ATW.")
        mapping[left] = right.upper()
    return mapping


def mapping_text(mapping: Dict[str, str]) -> str:
    return "\n".join(f"{k}={v}" for k, v in sorted(mapping.items()))


def reset_for_new_payroll(sig: str) -> None:
    if st.session_state.get("payroll_sig") == sig:
        return
    st.session_state.payroll_sig = sig
    st.session_state.corrections = {}
    st.session_state.overrides = {}
    st.session_state.report_bytes = None
    st.session_state.report_name = None


def missing_rows_df(lines: List[ParsedLine], corrections: Dict[str, str], mapping: Dict[str, str]) -> pd.DataFrame:
    rows = []
    for line in review_lines(lines):
        if line.nominal_code:
            continue
        key = missing_nominal_key(line)
        code = clean_nominal(corrections.get(key, ""))
        rows.append({
            "Key": key,
            "Staff Key": line.staff_key,
            "Row No": line.row_no,
            "Staff Name": line.staff_name,
            "Period": line.period,
            "Cost Type": "NIC" if line.is_nic else "WAGES",
            "Other Category": line.category,
            "Amount": round(abs(line.amount), 2),
            "Nominal": code,
            "Home": nominal_home(code, mapping) if code else "",
            "Description": line.description,
        })
    return pd.DataFrame(rows)


def correction_review_df(lines: List[ParsedLine], corrections: Dict[str, str], mapping: Dict[str, str]) -> pd.DataFrame:
    rows = []
    for line in review_lines(lines):
        key = missing_nominal_key(line)
        if key not in corrections:
            continue
        code = clean_nominal(corrections.get(key, ""))
        rows.append({
            "Key": key,
            "Staff Key": line.staff_key,
            "Row No": line.row_no,
            "Staff Name": line.staff_name,
            "Period": line.period,
            "Cost Type": "NIC" if line.is_nic else "WAGES",
            "Other Category": line.category,
            "Amount": round(abs(line.amount), 2),
            "Nominal": code,
            "Home": nominal_home(code, mapping) if code else "",
            "Description": line.description,
        })
    return pd.DataFrame(rows)


def save_nominals_from_df(df: pd.DataFrame, corrections: Dict[str, str]) -> List[str]:
    errors: List[str] = []
    for _, row in df.iterrows():
        raw = row.get("Nominal", "")
        code = clean_nominal(raw)
        if not code:
            continue
        ok, cleaned_or_error = validate_nominal(code)
        if not ok:
            errors.append(f"Row {row.get('Row No')}: {cleaned_or_error}")
            continue
        corrections[str(row["Key"])] = cleaned_or_error
    return errors


def staff_base_df(base: Dict[str, Dict[str, object]], mapping: Dict[str, str]) -> pd.DataFrame:
    rows = []
    for key, item in sorted(base.items(), key=lambda kv: str(kv[1].get("staff_name", "")).lower()):
        base_nominal = clean_nominal(item.get("base_nominal", ""))
        base_home = str(item.get("base_home", "")) or nominal_home(base_nominal, mapping)
        rows.append({
            "Staff Key": key,
            "Staff Name": item.get("staff_name", ""),
            "Base Nominal": base_nominal,
            "Base Home": base_home,
            "Source": item.get("source", ""),
            "Confidence": item.get("confidence", ""),
            "Employee Code": item.get("employee_code", ""),
            "Department": item.get("department", ""),
        })
    return pd.DataFrame(rows)


def save_staff_overrides_from_df(df: pd.DataFrame, overrides: Dict[str, str]) -> List[str]:
    errors: List[str] = []
    for _, row in df.iterrows():
        code = clean_nominal(row.get("Base Nominal", ""))
        if not code:
            continue
        ok, cleaned_or_error = validate_nominal(code)
        if not ok:
            errors.append(f"{row.get('Staff Name')}: {cleaned_or_error}")
            continue
        overrides[str(row["Staff Key"])] = cleaned_or_error
    return errors


st.title("Nominal OTHER Review")
st.caption("Upload payroll and staff master files, review missing nominals, then download the Excel report.")

if "mapping_text" not in st.session_state:
    st.session_state.mapping_text = mapping_text(DEFAULT_HOME_MAPPING)
if "corrections" not in st.session_state:
    st.session_state.corrections = {}
if "overrides" not in st.session_state:
    st.session_state.overrides = {}
if "report_bytes" not in st.session_state:
    st.session_state.report_bytes = None
if "report_name" not in st.session_state:
    st.session_state.report_name = None

with st.sidebar:
    st.subheader("Nominal mapping")
    st.session_state.mapping_text = st.text_area(
        "Prefix to home",
        value=st.session_state.mapping_text,
        height=150,
        help="One mapping per line, for example 2=CRE.",
    )
    st.caption("Default mapping: 2 CRE, 3 PET, 4 ATW, 5 BUN, 6 ALI, 7 HO.")

try:
    mapping = parse_mapping(st.session_state.mapping_text)
except ValueError as exc:
    st.error(str(exc))
    st.stop()

payroll_file = st.file_uploader("Nominal Update Report CSV/XLSX", type=["csv", "xlsx", "xlsm", "xltx", "xltm"])
master_file = st.file_uploader("Staff master XLSX", type=["xlsx", "xlsm", "xltx", "xltm"])

if not payroll_file:
    st.info("Upload the Nominal Update Report to begin.")
    st.stop()

reset_for_new_payroll(uploaded_signature(payroll_file))

with tempfile.TemporaryDirectory() as temp_dir_name:
    temp_dir = Path(temp_dir_name)
    payroll_path = save_uploaded_file(payroll_file, temp_dir, "payroll")
    master_path = save_uploaded_file(master_file, temp_dir, "staff_master") if master_file else None

    try:
        lines = read_lines(payroll_path)
        other_lines = review_lines(lines)
        if not other_lines:
            st.error("No OTHER or Employer NIC on OTHER lines were found in the payroll report.")
            st.stop()
    except Exception as exc:
        st.error(f"Could not read payroll report: {exc}")
        st.stop()

    apply_missing_nominals(lines, st.session_state.corrections)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Payroll lines parsed", f"{len(lines):,}")
    col2.metric("OTHER/NIC rows", f"{len(other_lines):,}")
    col3.metric("Staff with OTHER rows", f"{len({line.staff_key for line in other_lines}):,}")
    col4.metric("Saved missing nominals", f"{len(st.session_state.corrections):,}")

    st.divider()

    st.subheader("1. Missing nominal review")
    missing_df = missing_rows_df(lines, st.session_state.corrections, mapping)
    known_options = nominal_dropdown_options(lines, mapping)
    if known_options:
        st.caption("Known nominal suggestions: " + ", ".join(known_options[:40]))

    if not missing_df.empty:
        st.warning(f"{len(missing_df)} OTHER/NIC row(s) have blank nominal codes.")
        edited_missing = st.data_editor(
            missing_df,
            hide_index=True,
            use_container_width=True,
            disabled=["Key", "Staff Key", "Row No", "Staff Name", "Period", "Cost Type", "Other Category", "Amount", "Home", "Description"],
            column_config={
                "Nominal": st.column_config.TextColumn("Nominal", help="Enter the payroll nominal code, e.g. 504."),
                "Amount": st.column_config.NumberColumn("Amount", format="%.2f"),
            },
            key="missing_editor",
        )
        left, mid, right = st.columns([1.2, 1.6, 3])
        if left.button("Save nominal edits", type="primary"):
            errors = save_nominals_from_df(edited_missing, st.session_state.corrections)
            if errors:
                st.error("\n".join(errors))
            else:
                st.success("Missing nominal edits saved for this session.")
                st.rerun()
        if mid.button("Fill each staff from first entry"):
            errors = save_nominals_from_df(edited_missing, st.session_state.corrections)
            first_by_staff: Dict[str, str] = {}
            for _, row in edited_missing.iterrows():
                code = clean_nominal(row.get("Nominal", ""))
                if code and str(row["Staff Key"]) not in first_by_staff:
                    first_by_staff[str(row["Staff Key"])] = code
            for _, row in edited_missing.iterrows():
                staff = str(row["Staff Key"])
                if staff in first_by_staff:
                    st.session_state.corrections[str(row["Key"])] = first_by_staff[staff]
            if errors:
                st.error("\n".join(errors))
            else:
                st.success("Filled each staff member's missing rows from their first entered nominal.")
                st.rerun()
        st.stop()
    else:
        st.success("No blank OTHER/NIC nominal codes remain.")

    correction_df = correction_review_df(lines, st.session_state.corrections, mapping)
    if not correction_df.empty:
        with st.expander("Review saved missing nominal corrections", expanded=True):
            edited_corrections = st.data_editor(
                correction_df,
                hide_index=True,
                use_container_width=True,
                disabled=["Key", "Staff Key", "Row No", "Staff Name", "Period", "Cost Type", "Other Category", "Amount", "Home", "Description"],
                key="correction_editor",
            )
            if st.button("Update saved corrections"):
                errors = save_nominals_from_df(edited_corrections, st.session_state.corrections)
                if errors:
                    st.error("\n".join(errors))
                else:
                    st.success("Corrections updated.")
                    st.rerun()

    unknown_prefixes = sorted({
        line.nominal_code[0]
        for line in review_lines(lines)
        if line.nominal_code and line.nominal_code[0] not in mapping
    })
    if unknown_prefixes:
        st.error("These nominal prefixes are not mapped: " + ", ".join(unknown_prefixes))
        st.info("Add them in the sidebar mapping, for example 8=NEW.")
        st.stop()

    st.divider()

    st.subheader("2. Staff base review")
    try:
        staff_master = load_staff_master(master_path, mapping)
    except Exception as exc:
        st.error(f"Could not read staff master file: {exc}")
        st.stop()

    base = infer_staff_base(lines, mapping, staff_master, st.session_state.overrides)
    base_table = staff_base_df(base, mapping)
    missing_base = base_table[base_table["Base Home"].astype(str).str.strip() == ""] if not base_table.empty else pd.DataFrame()

    with st.expander("Review/edit staff base homes", expanded=not missing_base.empty):
        edited_base = st.data_editor(
            base_table,
            hide_index=True,
            use_container_width=True,
            disabled=["Staff Key", "Staff Name", "Base Home", "Source", "Confidence", "Employee Code", "Department"],
            column_config={
                "Base Nominal": st.column_config.TextColumn("Base Nominal", help="Enter a nominal such as 208, 404, 504 to override."),
            },
            key="base_editor",
        )
        if st.button("Save staff base overrides"):
            errors = save_staff_overrides_from_df(edited_base, st.session_state.overrides)
            if errors:
                st.error("\n".join(errors))
            else:
                st.success("Staff base overrides saved for this session.")
                st.rerun()

    if not missing_base.empty:
        st.error(f"{len(missing_base)} staff member(s) still need a base home. Enter Base Nominal overrides above.")
        st.stop()

    review_df, transfer_df, nominal_df, staff_df = build_review_rows(lines, base, mapping)
    unresolved = review_df[review_df["Status"].astype(str).str.startswith("REVIEW")] if not review_df.empty else pd.DataFrame()
    if not unresolved.empty:
        st.error(f"{len(unresolved)} review row(s) remain unresolved.")
        st.dataframe(unresolved, use_container_width=True, hide_index=True)
        st.stop()

    st.divider()
    st.subheader("3. Generate report")
    status_counts = review_df["Status"].value_counts().to_dict() if not review_df.empty else {}
    g1, g2, g3 = st.columns(3)
    g1.metric("Transfer rows", f"{int(status_counts.get('TRANSFER_TO_OTHER_HOME', 0)):,}")
    g2.metric("Same-home rows", f"{int(status_counts.get('SAME_HOME_OTHER_JOB_OR_SAME_HOME_OTHER', 0)):,}")
    g3.metric("Transfer total", f"\u00a3{float(transfer_df['Amount'].sum()) if not transfer_df.empty else 0.0:,.2f}")

    if st.button("Generate Excel report", type="primary"):
        output_path = temp_dir / default_output_name(payroll_file.name)
        write_output(
            output_path,
            review_df,
            transfer_df,
            nominal_df,
            staff_df,
            mapping,
            st.session_state.corrections,
            st.session_state.overrides,
        )
        st.session_state.report_bytes = output_path.read_bytes()
        st.session_state.report_name = output_path.name
        st.success("Report generated.")

    if st.session_state.report_bytes:
        st.download_button(
            "Download Excel report",
            data=st.session_state.report_bytes,
            file_name=st.session_state.report_name or default_output_name(payroll_file.name),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
