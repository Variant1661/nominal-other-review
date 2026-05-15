from __future__ import annotations

import hashlib
import tempfile
from pathlib import Path
from typing import Dict, List, Tuple

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
    validate_nominal,
    write_output,
)


APP_DIR = Path(__file__).resolve().parent
LOGO_PATH = APP_DIR / "assets" / "company_logo.gif"

st.set_page_config(
    page_title="Nominal OTHER Review",
    page_icon=str(LOGO_PATH) if LOGO_PATH.exists() else None,
    layout="wide",
)

st.markdown(
    """
    <style>
    :root {
        --app-bg: #f4f7fb;
        --panel: #ffffff;
        --ink: #172033;
        --muted: #667085;
        --line: #d8e1ef;
        --brand: #1f5fbf;
        --brand-dark: #164a99;
        --green: #12805c;
        --amber: #b7791f;
        --red: #c2413a;
    }
    .stApp {background: var(--app-bg);}
    .block-container {padding-top: 1.2rem; max-width: 1420px;}
    h1, h2, h3 {letter-spacing: 0;}
    div[data-testid="stMetric"] {
        background: var(--panel);
        border: 1px solid var(--line);
        border-radius: 10px;
        padding: 14px 16px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
    }
    div[data-testid="stMetricValue"] {font-size: 1.45rem;}
    .stButton button, .stDownloadButton button {
        border-radius: 7px;
        font-weight: 700;
        min-height: 2.5rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background: #eaf0f8;
        border-radius: 10px;
        padding: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 42px;
        border-radius: 8px;
        padding: 0 18px;
        font-weight: 700;
    }
    .app-hero {
        background: linear-gradient(135deg, #ffffff 0%, #eef5ff 100%);
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 18px 22px;
        box-shadow: 0 14px 34px rgba(15, 23, 42, 0.06);
        margin-bottom: 16px;
    }
    .hero-title {font-size: 30px; font-weight: 800; color: var(--ink); margin: 0;}
    .hero-subtitle {font-size: 14px; color: var(--muted); margin-top: 4px;}
    .brand-wordmark {font-size: 14px; font-weight: 800; color: var(--ink);}
    .status-ribbon {
        border: 1px solid var(--line);
        border-radius: 10px;
        background: #ffffff;
        padding: 12px 16px;
        margin: 8px 0 16px 0;
    }
    .pill {
        display: inline-block;
        padding: 5px 10px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 800;
        margin-right: 6px;
        border: 1px solid transparent;
    }
    .pill-ready {background: #e8f7f0; color: var(--green); border-color: #b7e4cf;}
    .pill-action {background: #fff7e6; color: var(--amber); border-color: #f6d79b;}
    .pill-blocked {background: #fdeceb; color: var(--red); border-color: #f4b6b2;}
    .section-card {
        background: var(--panel);
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 16px 18px;
        box-shadow: 0 10px 26px rgba(15, 23, 42, 0.04);
        margin-bottom: 14px;
    }
    .section-title {font-size: 18px; font-weight: 800; color: var(--ink); margin-bottom: 4px;}
    .section-help {font-size: 13px; color: var(--muted); margin-bottom: 10px;}
    .small-muted {color: var(--muted); font-size: 12px;}
    </style>
    """,
    unsafe_allow_html=True,
)


def uploaded_signature(uploaded_file: object) -> str:
    return hashlib.sha256(uploaded_file.getvalue()).hexdigest()


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


def missing_staff_df(lines: List[ParsedLine], corrections: Dict[str, str], mapping: Dict[str, str]) -> pd.DataFrame:
    missing = missing_rows_df(lines, corrections, mapping)
    if missing.empty:
        return pd.DataFrame(columns=["Staff Key", "Staff Name", "Rows", "Categories", "Cost Types", "Total Amount", "Nominal", "Home", "Row Nos"])

    rows = []
    for staff, group in missing.groupby("Staff Key", sort=True):
        codes = [clean_nominal(v) for v in group["Nominal"].tolist() if clean_nominal(v)]
        code = codes[0] if codes else ""
        rows.append({
            "Staff Key": staff,
            "Staff Name": group["Staff Name"].iloc[0],
            "Rows": int(len(group)),
            "Categories": ", ".join(sorted(set(group["Other Category"].astype(str)))),
            "Cost Types": ", ".join(sorted(set(group["Cost Type"].astype(str)))),
            "Total Amount": round(float(group["Amount"].sum()), 2),
            "Nominal": code,
            "Home": nominal_home(code, mapping) if code else "",
            "Row Nos": ", ".join(str(v) for v in sorted(group["Row No"].tolist())),
        })
    return pd.DataFrame(rows).sort_values("Staff Name")


def correction_rows_df(lines: List[ParsedLine], corrections: Dict[str, str], mapping: Dict[str, str]) -> pd.DataFrame:
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


def correction_staff_df(lines: List[ParsedLine], corrections: Dict[str, str], mapping: Dict[str, str]) -> pd.DataFrame:
    details = correction_rows_df(lines, corrections, mapping)
    if details.empty:
        return pd.DataFrame(columns=["Staff Key", "Staff Name", "Rows", "Categories", "Cost Types", "Total Amount", "Nominal", "Home", "Row Nos"])

    rows = []
    for staff, group in details.groupby("Staff Key", sort=True):
        codes = [clean_nominal(v) for v in group["Nominal"].tolist() if clean_nominal(v)]
        code = codes[0] if codes else ""
        rows.append({
            "Staff Key": staff,
            "Staff Name": group["Staff Name"].iloc[0],
            "Rows": int(len(group)),
            "Categories": ", ".join(sorted(set(group["Other Category"].astype(str)))),
            "Cost Types": ", ".join(sorted(set(group["Cost Type"].astype(str)))),
            "Total Amount": round(float(group["Amount"].sum()), 2),
            "Nominal": code,
            "Home": nominal_home(code, mapping) if code else "",
            "Row Nos": ", ".join(str(v) for v in sorted(group["Row No"].tolist())),
        })
    return pd.DataFrame(rows).sort_values("Staff Name")


def save_missing_staff_nominals(df: pd.DataFrame, lines: List[ParsedLine], corrections: Dict[str, str]) -> List[str]:
    errors: List[str] = []
    code_by_staff: Dict[str, str] = {}

    for _, row in df.iterrows():
        code = clean_nominal(row.get("Nominal", ""))
        if not code:
            continue
        ok, cleaned_or_error = validate_nominal(code)
        if not ok:
            errors.append(f"{row.get('Staff Name')}: {cleaned_or_error}")
            continue
        code_by_staff[str(row["Staff Key"])] = cleaned_or_error

    if errors:
        return errors

    for line in review_lines(lines):
        if line.nominal_code:
            continue
        code = code_by_staff.get(line.staff_key)
        if code:
            corrections[missing_nominal_key(line)] = code
    return []


def save_correction_staff_nominals(df: pd.DataFrame, lines: List[ParsedLine], corrections: Dict[str, str]) -> List[str]:
    errors: List[str] = []
    code_by_staff: Dict[str, str] = {}

    for _, row in df.iterrows():
        code = clean_nominal(row.get("Nominal", ""))
        if not code:
            continue
        ok, cleaned_or_error = validate_nominal(code)
        if not ok:
            errors.append(f"{row.get('Staff Name')}: {cleaned_or_error}")
            continue
        code_by_staff[str(row["Staff Key"])] = cleaned_or_error

    if errors:
        return errors

    for line in review_lines(lines):
        key = missing_nominal_key(line)
        if key in corrections and line.staff_key in code_by_staff:
            corrections[key] = code_by_staff[line.staff_key]
    return []


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


def render_header() -> None:
    logo_col, title_col = st.columns([0.12, 0.88], vertical_alignment="center")
    with logo_col:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=72)
        else:
            st.markdown("<div class='brand-wordmark'>Apex<br>Care Homes</div>", unsafe_allow_html=True)
    with title_col:
        st.markdown(
            """
            <div class="app-hero">
                <div class="hero-title">Nominal OTHER Review</div>
                <div class="hero-subtitle">Apex Care Homes payroll transfer review workspace</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_status_ribbon(blockers: List[str], ready: bool) -> None:
    if ready:
        status = "<span class='pill pill-ready'>READY TO GENERATE</span>"
        message = "All required checks are complete."
    else:
        status = "<span class='pill pill-action'>ACTION NEEDED</span>"
        message = "Complete the work queue items before generating the Excel report."
    blocker_text = "" if ready else " ".join(f"<span class='pill pill-blocked'>{item}</span>" for item in blockers)
    st.markdown(f"<div class='status-ribbon'>{status} {message}<br>{blocker_text}</div>", unsafe_allow_html=True)


def render_section(title: str, help_text: str) -> None:
    st.markdown(
        f"<div class='section-card'><div class='section-title'>{title}</div><div class='section-help'>{help_text}</div></div>",
        unsafe_allow_html=True,
    )


def dataframe_height(df: pd.DataFrame, minimum: int = 220, maximum: int = 520) -> int:
    if df.empty:
        return minimum
    return max(minimum, min(maximum, 80 + len(df) * 30))


def render_dashboard(
    lines: List[ParsedLine],
    review_df: pd.DataFrame,
    transfer_df: pd.DataFrame,
    missing_staff: pd.DataFrame,
    correction_staff: pd.DataFrame,
    missing_base: pd.DataFrame,
    blockers: List[str],
) -> None:
    status_counts = review_df["Status"].value_counts().to_dict() if not review_df.empty else {}
    transfer_total = float(transfer_df["Amount"].sum()) if not transfer_df.empty else 0.0
    other_lines = review_lines(lines)

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Payroll lines", f"{len(lines):,}")
    m2.metric("OTHER/NIC rows", f"{len(other_lines):,}")
    m3.metric("Staff in review", f"{len({line.staff_key for line in other_lines}):,}")
    m4.metric("Missing staff queue", f"{len(missing_staff):,}")
    m5.metric("Transfer total", f"\u00a3{transfer_total:,.2f}")

    st.markdown("#### Live dashboard")
    left, right = st.columns([0.52, 0.48])
    with left:
        if not review_df.empty:
            status_chart = pd.DataFrame(
                {"Rows": pd.Series(status_counts)}
            ).sort_values("Rows", ascending=False)
            st.bar_chart(status_chart, use_container_width=True)
        else:
            st.info("Review rows will appear after upload.")
    with right:
        if not review_df.empty:
            top_staff = (
                review_df.groupby("Staff Name")["Amount"]
                .sum()
                .sort_values(ascending=False)
                .head(10)
            )
            st.bar_chart(top_staff, use_container_width=True)
        else:
            st.info("Staff totals will appear after upload.")

    c1, c2 = st.columns([0.52, 0.48])
    with c1:
        st.markdown("#### Transfer summary preview")
        if transfer_df.empty:
            st.info("No transfer rows yet.")
        else:
            st.dataframe(transfer_df, use_container_width=True, hide_index=True, height=min(380, dataframe_height(transfer_df)))
    with c2:
        st.markdown("#### Work queue")
        queue = pd.DataFrame([
            {"Area": "Missing nominals", "Open items": len(missing_staff), "Status": "Clear" if missing_staff.empty else "Needs action"},
            {"Area": "Saved corrections", "Open items": len(correction_staff), "Status": "Reviewable" if not correction_staff.empty else "None"},
            {"Area": "Staff base homes", "Open items": len(missing_base), "Status": "Clear" if missing_base.empty else "Needs action"},
            {"Area": "Final blockers", "Open items": len(blockers), "Status": "Clear" if not blockers else "Needs action"},
        ])
        st.dataframe(queue, use_container_width=True, hide_index=True, height=210)


render_header()

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
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=96)
    st.subheader("Workspace controls")
    st.session_state.mapping_text = st.text_area(
        "Nominal prefix mapping",
        value=st.session_state.mapping_text,
        height=150,
        help="One mapping per line, for example 2=CRE.",
    )
    st.caption("Default mapping: 2 CRE, 3 PET, 4 ATW, 5 BUN, 6 ALI, 7 HO.")
    st.divider()
    st.caption("Files are processed in the active Streamlit session. Download the generated workbook before closing the page.")

try:
    mapping = parse_mapping(st.session_state.mapping_text)
except ValueError as exc:
    st.error(str(exc))
    st.stop()

upload_left, upload_right = st.columns(2)
with upload_left:
    payroll_file = st.file_uploader("Nominal Update Report CSV/XLSX", type=["csv", "xlsx", "xlsm", "xltx", "xltm"])
with upload_right:
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

    try:
        staff_master = load_staff_master(master_path, mapping)
    except Exception as exc:
        st.error(f"Could not read staff master file: {exc}")
        st.stop()

    base = infer_staff_base(lines, mapping, staff_master, st.session_state.overrides)
    base_table = staff_base_df(base, mapping)
    missing_base = base_table[base_table["Base Home"].astype(str).str.strip() == ""] if not base_table.empty else pd.DataFrame()
    review_df, transfer_df, nominal_df, staff_df = build_review_rows(lines, base, mapping)
    missing_staff = missing_staff_df(lines, st.session_state.corrections, mapping)
    missing_detail = missing_rows_df(lines, st.session_state.corrections, mapping)
    correction_staff = correction_staff_df(lines, st.session_state.corrections, mapping)
    correction_detail = correction_rows_df(lines, st.session_state.corrections, mapping)

    unknown_prefixes = sorted({
        line.nominal_code[0]
        for line in review_lines(lines)
        if line.nominal_code and line.nominal_code[0] not in mapping
    })
    unresolved = review_df[review_df["Status"].astype(str).str.startswith("REVIEW")] if not review_df.empty else pd.DataFrame()

    blockers: List[str] = []
    if not missing_staff.empty:
        blockers.append(f"{len(missing_staff)} staff need nominal")
    if unknown_prefixes:
        blockers.append("Unmapped prefix " + ", ".join(unknown_prefixes))
    if not missing_base.empty:
        blockers.append(f"{len(missing_base)} base homes")
    if not unresolved.empty and missing_staff.empty and missing_base.empty:
        blockers.append(f"{len(unresolved)} review rows")
    ready = not blockers and unresolved.empty

    render_status_ribbon(blockers, ready)

    dashboard_tab, missing_tab, staff_tab, detail_tab, generate_tab = st.tabs(
        ["Dashboard", "Missing Nominals", "Staff Base", "Review Detail", "Generate"]
    )

    with dashboard_tab:
        render_dashboard(lines, review_df, transfer_df, missing_staff, correction_staff, missing_base, blockers)

    with missing_tab:
        render_section(
            "Missing nominal work queue",
            "Enter one nominal per staff member. The app automatically applies that code to all missing OTHER/NIC rows for that staff.",
        )
        known_options = nominal_dropdown_options(lines, mapping)
        if known_options:
            st.caption("Known nominal suggestions: " + ", ".join(known_options[:45]))

        if missing_staff.empty:
            st.success("No blank OTHER/NIC nominal codes remain.")
        else:
            edited_missing_staff = st.data_editor(
                missing_staff,
                hide_index=True,
                use_container_width=True,
                height=dataframe_height(missing_staff),
                disabled=["Staff Key", "Staff Name", "Rows", "Categories", "Cost Types", "Total Amount", "Home", "Row Nos"],
                column_config={
                    "Nominal": st.column_config.TextColumn("Nominal", help="Enter once; it repeats to all missing rows for this staff."),
                    "Total Amount": st.column_config.NumberColumn("Total Amount", format="%.2f"),
                },
                key="missing_staff_editor",
            )
            if st.button("Save and apply to all rows for each staff", type="primary"):
                errors = save_missing_staff_nominals(edited_missing_staff, lines, st.session_state.corrections)
                if errors:
                    st.error("\n".join(errors))
                else:
                    st.success("Saved. Each entered nominal was repeated to all missing OTHER/NIC rows for that staff.")
                    st.rerun()

            with st.expander("Line-level missing nominal detail"):
                st.dataframe(missing_detail, use_container_width=True, hide_index=True, height=dataframe_height(missing_detail))

        if not correction_staff.empty:
            st.markdown("#### Saved manual corrections")
            edited_correction_staff = st.data_editor(
                correction_staff,
                hide_index=True,
                use_container_width=True,
                height=dataframe_height(correction_staff, minimum=180),
                disabled=["Staff Key", "Staff Name", "Rows", "Categories", "Cost Types", "Total Amount", "Home", "Row Nos"],
                column_config={
                    "Nominal": st.column_config.TextColumn("Nominal", help="Changing this repeats to all saved correction rows for this staff."),
                    "Total Amount": st.column_config.NumberColumn("Total Amount", format="%.2f"),
                },
                key="correction_staff_editor",
            )
            if st.button("Update saved staff corrections"):
                errors = save_correction_staff_nominals(edited_correction_staff, lines, st.session_state.corrections)
                if errors:
                    st.error("\n".join(errors))
                else:
                    st.success("Saved corrections updated by staff.")
                    st.rerun()
            with st.expander("Line-level saved correction detail"):
                st.dataframe(correction_detail, use_container_width=True, hide_index=True, height=dataframe_height(correction_detail))

    with staff_tab:
        render_section(
            "Staff base review",
            "Review inferred or staff-master base homes. Add a Base Nominal only where an override is needed.",
        )
        edited_base = st.data_editor(
            base_table,
            hide_index=True,
            use_container_width=True,
            height=dataframe_height(base_table),
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
            st.error(f"{len(missing_base)} staff member(s) still need a base home.")

    with detail_tab:
        render_section("Review detail", "Inspect the rows that will feed the final workbook.")
        view = st.segmented_control("View", ["OTHER Review", "Transfer Summary", "Nominal Summary"], default="OTHER Review")
        if view == "Transfer Summary":
            st.dataframe(transfer_df, use_container_width=True, hide_index=True, height=dataframe_height(transfer_df))
        elif view == "Nominal Summary":
            st.dataframe(nominal_df, use_container_width=True, hide_index=True, height=dataframe_height(nominal_df))
        else:
            st.dataframe(review_df, use_container_width=True, hide_index=True, height=dataframe_height(review_df, maximum=640))

    with generate_tab:
        render_section("Generate workbook", "When the work queue is clear, generate and download the Excel review report.")
        if unknown_prefixes:
            st.error("These nominal prefixes are not mapped: " + ", ".join(unknown_prefixes))
            st.info("Add them in the sidebar mapping, for example 8=NEW.")
        if not ready:
            st.warning("The report is not ready yet. Clear the blockers shown at the top of the page.")
        else:
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
