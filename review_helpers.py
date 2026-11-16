"""Table-building and save helpers for the Streamlit app (no Streamlit import, so testable)."""
from __future__ import annotations

from typing import Dict, List

import pandas as pd

from nominal_logic import (
    ParsedLine,
    clean_nominal,
    missing_nominal_key,
    nominal_home,
    review_lines,
    validate_nominal,
)

STAFF_COLUMNS = ["Staff Key", "Staff Name", "Rows", "Categories", "Cost Types", "Total Amount", "Nominal", "Home", "Row Nos"]


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


def _row(line: ParsedLine, code: str, mapping: Dict[str, str]) -> Dict[str, object]:
    return {
        "Key": missing_nominal_key(line),
        "Staff Key": line.staff_key,
        "Row No": line.row_no,
        "Staff Name": line.staff_name,
        "Period": line.period,
        "Cost Type": "NIC" if line.is_nic else "WAGES",
        "Other Category": line.category,
        "Amount": round(line.amount, 2),
        "Nominal": code,
        "Home": nominal_home(code, mapping) if code else "",
        "Description": line.description,
    }


def missing_rows_df(lines: List[ParsedLine], corrections: Dict[str, str], mapping: Dict[str, str]) -> pd.DataFrame:
    rows = []
    for line in review_lines(lines):
        if line.nominal_code:
            continue
        rows.append(_row(line, clean_nominal(corrections.get(missing_nominal_key(line), "")), mapping))
    return pd.DataFrame(rows)


def correction_rows_df(lines: List[ParsedLine], corrections: Dict[str, str], mapping: Dict[str, str]) -> pd.DataFrame:
    rows = []
    for line in review_lines(lines):
        key = missing_nominal_key(line)
        if key in corrections:
            rows.append(_row(line, clean_nominal(corrections[key]), mapping))
    return pd.DataFrame(rows)


def _staff_rollup(details: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
    if details.empty:
        return pd.DataFrame(columns=STAFF_COLUMNS)
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


def missing_staff_df(lines: List[ParsedLine], corrections: Dict[str, str], mapping: Dict[str, str]) -> pd.DataFrame:
    return _staff_rollup(missing_rows_df(lines, corrections, mapping), mapping)


def correction_staff_df(lines: List[ParsedLine], corrections: Dict[str, str], mapping: Dict[str, str]) -> pd.DataFrame:
    return _staff_rollup(correction_rows_df(lines, corrections, mapping), mapping)


def _staff_codes(df: pd.DataFrame, original: pd.DataFrame | None = None):
    """Validated code per staff key. With `original`, only staff whose Nominal was
    edited are returned, so untouched rows keep their own saved values."""
    errors: List[str] = []
    codes: Dict[str, str] = {}
    before = {}
    if original is not None and not original.empty:
        before = {str(r["Staff Key"]): clean_nominal(r.get("Nominal", "")) for _, r in original.iterrows()}
    for _, row in df.iterrows():
        code = clean_nominal(row.get("Nominal", ""))
        key = str(row["Staff Key"])
        if not code or (original is not None and before.get(key, "") == code):
            continue
        ok, cleaned_or_error = validate_nominal(code)
        if not ok:
            errors.append(f"{row.get('Staff Name')}: {cleaned_or_error}")
            continue
        codes[key] = cleaned_or_error
    return codes, errors


def save_missing_staff_nominals(df: pd.DataFrame, lines: List[ParsedLine], corrections: Dict[str, str]) -> List[str]:
    code_by_staff, errors = _staff_codes(df)
    if errors:
        return errors
    for line in review_lines(lines):
        if line.nominal_code:
            continue
        code = code_by_staff.get(line.staff_key)
        if code:
            corrections[missing_nominal_key(line)] = code
    return []


def save_correction_staff_nominals(
    df: pd.DataFrame, lines: List[ParsedLine], corrections: Dict[str, str], original: pd.DataFrame
) -> List[str]:
    """Apply an edited staff-level code to that staff's saved rows - only for staff
    whose code was actually changed in the table."""
    code_by_staff, errors = _staff_codes(df, original)
    if errors:
        return errors
    for line in review_lines(lines):
        key = missing_nominal_key(line)
        if key in corrections and line.staff_key in code_by_staff:
            corrections[key] = code_by_staff[line.staff_key]
    return []


def save_review_nominal_overrides(
    df: pd.DataFrame,
    corrections: Dict[str, str],
    correction_key_by_row: Dict[int, str],
    original_nominal_by_row: Dict[int, str],
) -> List[str]:
    errors: List[str] = []
    updates: Dict[str, str] = {}
    for _, row in df.iterrows():
        row_no_text = clean_nominal(row.get("Row No", ""))
        row_no = int(row_no_text) if row_no_text else 0
        key = correction_key_by_row.get(row_no, "")
        if not key:
            continue
        code = clean_nominal(row.get("Payroll Nominal", ""))
        original_code = clean_nominal(original_nominal_by_row.get(row_no, ""))
        if code == original_code:
            continue
        ok, cleaned_or_error = validate_nominal(code)
        if not ok:
            errors.append(f"{row.get('Staff Name')}: {cleaned_or_error}")
            continue
        updates[key] = cleaned_or_error
    if errors:
        return errors
    corrections.update(updates)
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
            "Note": item.get("note", ""),
        })
    return pd.DataFrame(rows)


def save_staff_overrides_from_df(df: pd.DataFrame, overrides: Dict[str, str], original: pd.DataFrame) -> List[str]:
    """Save only edited Base Nominal cells. A cell cleared by the user removes that
    staff member's override; untouched rows are never turned into manual overrides."""
    before = {str(r["Staff Key"]): clean_nominal(r.get("Base Nominal", "")) for _, r in original.iterrows()}
    errors: List[str] = []
    updates: Dict[str, str] = {}
    removals: List[str] = []
    for _, row in df.iterrows():
        key = str(row["Staff Key"])
        code = clean_nominal(row.get("Base Nominal", ""))
        if code == before.get(key, ""):
            continue
        if not code:
            removals.append(key)
            continue
        ok, cleaned_or_error = validate_nominal(code)
        if not ok:
            errors.append(f"{row.get('Staff Name')}: {cleaned_or_error}")
            continue
        updates[key] = cleaned_or_error
    if errors:
        return errors
    overrides.update(updates)
    for key in removals:
        overrides.pop(key, None)
    return []


def unusable_base_homes(base_table: pd.DataFrame) -> pd.DataFrame:
    """Staff whose base home is blank or UNKNOWN (an unmapped nominal prefix)."""
    if base_table.empty:
        return pd.DataFrame()
    home = base_table["Base Home"].astype(str).str.strip()
    return base_table[(home == "") | (home.str.upper() == "UNKNOWN")]
