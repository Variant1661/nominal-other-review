"""Department costing: re-allocate payroll cost between homes after wages are posted.

Two rule types, both driven by editable weights that are normalised to 100%:

* Pool rules  - every payroll line that matches (paying home + department suffix
  and/or job role) goes into one pool; the pool total is re-split between the
  listed homes by percentage (e.g. a kitchen shared by Crescent and Peter's).
* Person rules - one named person's total cost is re-split between homes by
  days or percentage (e.g. 2 days / 2 days / 1 day -> 40 / 40 / 20).
  A person rule always wins over a pool rule for that person.

The journal is the net movement needed so each home carries its share:
target (pool total x share) minus what is currently posted to that home.
Amounts are signed (debit - credit) so reversals reduce totals.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import pandas as pd
from openpyxl import load_workbook

from nominal_logic import (
    PAYROLL_LINE_RE,
    clean_amount,
    clean_nominal,
    clean_text,
    detect_columns,
    find_header_row,
    read_csv_rows,
    read_excel_rows,
    staff_key,
    staff_token_key,
)

DEFAULT_DEPARTMENTS: Dict[str, str] = {
    "04": "Care (Support Worker/Carer/Senior Carer)",
    "05": "Kitchen",
    "06": "Domestic (Housekeeping)",
    "08": "Nurse",
    "09": "Activities Coordinator",
    "10": "Laundry",
    "12": "Supernumerary (Nurses)",
}

MASTER_HOME_TO_CODE = {
    "crescent nursing home": "CRE",
    "crescent": "CRE",
    "peter s place": "PET",
    "peters place": "PET",
    "atwell house": "ATW",
    "bunyan lodge": "BUN",
    "alicia nursing home": "ALI",
    "head office": "HO",
}

TOLERANCE = 0.005


@dataclass
class CostLine:
    row_no: int
    nominal: str
    prefix: str
    suffix: str
    home: str
    amount: float
    cost_type: str  # WAGES / NIC / PENSION
    category: str
    staff_name: str
    person_key: str
    period: str
    description: str


@dataclass
class Rule:
    name: str
    kind: str  # "pool" or "person"
    weights: Dict[str, float]
    suffixes: List[str] = field(default_factory=list)
    job_roles: List[str] = field(default_factory=list)
    paying_homes: List[str] = field(default_factory=list)
    person: str = ""  # person rules: display name as typed/selected
    note: str = ""


# ---------------------------------------------------------------- reading --

def cost_type_of(category: str, is_nic: bool) -> str:
    if is_nic:
        return "NIC"
    if "PENSION" in category.upper():
        return "PENSION"
    return "WAGES"


def read_cost_lines(path: Path, mapping: Dict[str, str]) -> Tuple[List[CostLine], List[Dict[str, Any]]]:
    """All parseable payroll lines. Lines with no Code are returned separately
    (never inherit the nominal of the row above)."""
    rows = read_csv_rows(path) if path.suffix.lower() == ".csv" else read_excel_rows(path)
    header_idx = find_header_row(rows)
    cols = detect_columns(rows[header_idx])
    lines: List[CostLine] = []
    no_code: List[Dict[str, Any]] = []

    for zero_idx, row in enumerate(rows[header_idx + 1:], start=header_idx + 1):
        def get(name: str) -> Any:
            col = cols.get(name)
            return "" if col is None or col >= len(row) else row[col]

        desc = clean_text(get("description"))
        match = PAYROLL_LINE_RE.match(desc)
        if not match:
            continue
        debit, credit = clean_amount(get("debit")), clean_amount(get("credit"))
        amount = round(debit - abs(credit), 2)
        name = clean_text(match.group("name"))
        category = clean_text(match.group("category")).upper()
        is_nic = bool(match.group("nic"))
        nominal = clean_nominal(get("code"))
        if not nominal:
            no_code.append({"Row No": zero_idx + 1, "Staff Name": name, "Category": category,
                            "Cost Type": cost_type_of(category, is_nic), "Amount": amount,
                            "Issue": "No nominal code on this row - not allocated"})
            continue
        prefix = nominal[0]
        lines.append(CostLine(
            row_no=zero_idx + 1, nominal=nominal, prefix=prefix, suffix=nominal[1:],
            home=mapping.get(prefix, "UNKNOWN"), amount=amount,
            cost_type=cost_type_of(category, is_nic), category=category,
            staff_name=name, person_key=staff_token_key(name),
            period=clean_text(match.group("period")), description=desc,
        ))
    return lines, no_code


def load_master_roles(path: Optional[Path]) -> Tuple[Dict[str, List[Dict[str, str]]], pd.DataFrame]:
    """person_key -> [{name, home, job_role}] (a list so duplicates are visible)."""
    if not path:
        return {}, pd.DataFrame(columns=["Staff Name", "Home", "Job Role"])
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()
    header = [staff_key(clean_text(v)) for v in rows[0]]

    def col(*names: str) -> int:
        for n in names:
            if n in header:
                return header.index(n)
        raise ValueError(f"Staff master is missing a '{names[0]}' column.")

    i_name, i_home, i_role = col("full name"), col("home", "department"), col("job role")
    i_status = header.index("employment status") if "employment status" in header else None
    index: Dict[str, List[Dict[str, str]]] = defaultdict(list)
    flat = []
    for r in rows[1:]:
        if i_name >= len(r) or not clean_text(r[i_name]):
            continue
        if i_status is not None and i_status < len(r) and clean_text(r[i_status]).lower() not in ("", "active"):
            continue
        name = clean_text(r[i_name])
        home_text = staff_key(clean_text(r[i_home])) if i_home < len(r) else ""
        item = {
            "name": name,
            "home": MASTER_HOME_TO_CODE.get(home_text, home_text.upper()),
            "job_role": clean_text(r[i_role]) if i_role < len(r) else "",
        }
        index[staff_token_key(name)].append(item)
        flat.append({"Staff Name": name, "Home": item["home"], "Job Role": item["job_role"]})
    return dict(index), pd.DataFrame(flat)


# ---------------------------------------------------------------- rules ----

def normalise_weights(weights: Dict[str, float]) -> Dict[str, float]:
    clean = {h: float(w) for h, w in weights.items() if w is not None and float(w) > 0}
    total = sum(clean.values())
    if total <= 0:
        raise ValueError("Enter at least one positive share (percent or days).")
    return {h: w / total for h, w in clean.items()}


def rule_from_dict(d: Dict[str, Any]) -> Rule:
    split = lambda v: [x.strip() for x in str(v or "").split(",") if x.strip()]
    return Rule(
        name=str(d.get("name", "")).strip(), kind=str(d.get("kind", "pool")),
        weights={k: float(v) for k, v in (d.get("weights") or {}).items()},
        suffixes=[s.zfill(2) for s in split(",".join(d["suffixes"]) if isinstance(d.get("suffixes"), list) else d.get("suffixes"))],
        job_roles=split(",".join(d["job_roles"]) if isinstance(d.get("job_roles"), list) else d.get("job_roles")),
        paying_homes=[h.upper() for h in split(",".join(d["paying_homes"]) if isinstance(d.get("paying_homes"), list) else d.get("paying_homes"))],
        person=str(d.get("person", "")).strip(), note=str(d.get("note", "")),
    )


def rules_to_json(rules: List[Rule]) -> str:
    return json.dumps([r.__dict__ for r in rules], indent=2)


def rules_from_json(text: str) -> List[Rule]:
    return [rule_from_dict(d) for d in json.loads(text)]


def validate_rules(rules: List[Rule], mapping: Dict[str, str]) -> List[str]:
    errors: List[str] = []
    known = set(mapping.values())
    seen = set()
    for r in rules:
        label = r.name or "(unnamed rule)"
        if not r.name:
            errors.append("A rule has no name.")
        if r.name in seen:
            errors.append(f"Duplicate rule name: {r.name}")
        seen.add(r.name)
        try:
            normalise_weights(r.weights)
        except ValueError as exc:
            errors.append(f"{label}: {exc}")
        bad = [h for h in r.weights if h not in known]
        if bad:
            errors.append(f"{label}: unknown home(s) {', '.join(bad)}")
        if r.kind == "pool" and not (r.suffixes or r.job_roles):
            errors.append(f"{label}: pool rules need a department suffix and/or a job role.")
        if r.kind == "pool" and not r.paying_homes:
            errors.append(f"{label}: pool rules need at least one paying home.")
        if r.kind == "person" and not r.person:
            errors.append(f"{label}: person rules need a staff name.")
    return errors


# ------------------------------------------------------------ allocation ---

def _tokens(name: str) -> List[str]:
    return [t for t in staff_key(name).split() if t]


def resolve_staff(payroll_name: str, master: Dict[str, List[Dict[str, str]]]) -> Optional[Dict[str, str]]:
    """Match a payroll name ("Surname Forename X", initials trailing) to ONE master record.

    Rules: exact word-set match wins; otherwise every payroll word longer than one
    letter must appear in the master name (initials only break ties).
    More than one candidate (or none) returns None so it is flagged, never guessed.
    """
    key = staff_token_key(payroll_name)
    exact = master.get(key, [])
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        return None
    words = [t for t in _tokens(payroll_name) if len(t) > 1]
    initials = [t for t in _tokens(payroll_name) if len(t) == 1]
    if len(words) < 2:
        return None
    hits = [e for entries in master.values() for e in entries
            if all(w in _tokens(e["name"]) for w in words)]
    if len(hits) > 1 and initials:  # initials only break ties; master often omits middle initials
        def fits(e: Dict[str, str]) -> bool:
            rest = _tokens(e["name"])
            for w in words:
                rest.remove(w)
            return all(any(r.startswith(i) for r in rest) for i in initials)
        hits = [e for e in hits if fits(e)]
    return hits[0] if len(hits) == 1 else None


def assign_rules(lines: List[CostLine], rules: List[Rule], master: Dict[str, List[Dict[str, str]]]) -> Tuple[Dict[int, str], List[Dict[str, Any]]]:
    """index of line -> rule name. Person rules beat pool rules; first matching pool wins.
    Returns assignments plus warnings about people/rules that could not be resolved."""
    warnings: List[Dict[str, Any]] = []
    person_rules = {staff_token_key(r.person): r for r in rules if r.kind == "person"}
    resolved = {k: resolve_staff(n, master) for k, n in {ln.person_key: ln.staff_name for ln in lines}.items()}
    pool_rules = [r for r in rules if r.kind == "pool"]
    assigned: Dict[int, str] = {}
    matched_person = set()
    for i, ln in enumerate(lines):
        rec = resolved.get(ln.person_key)
        pr = person_rules.get(ln.person_key) or (person_rules.get(staff_token_key(rec["name"])) if rec else None)
        if pr:
            assigned[i] = pr.name
            matched_person.add(pr.name)
            continue
        role = rec["job_role"] if rec else ""
        for r in pool_rules:
            if ln.home not in r.paying_homes:
                continue
            if r.suffixes and ln.suffix not in r.suffixes:
                continue
            if r.job_roles:
                if not role:
                    continue
                if role.lower() not in [x.lower() for x in r.job_roles]:
                    continue
            assigned[i] = r.name
            break
    for r in rules:
        if r.kind == "person" and r.name not in matched_person:
            warnings.append({"Rule": r.name, "Issue": f"No payroll lines found for '{r.person}' - check the spelling."})
    # pool rules keyed on job role need master matching; say how many staff could not be matched
    if any(r.job_roles for r in pool_rules):
        miss = sorted({ln.staff_name for ln in lines if not resolved.get(ln.person_key)})
        if miss:
            warnings.append({"Rule": "(job-role rules)", "Issue": f"{len(miss)} payroll name(s) not matched to exactly one staff-master record, so job-role rules skip them: " + "; ".join(miss[:25]) + ("..." if len(miss) > 25 else "")})
    return assigned, warnings


def _net_transfers(actual: Dict[str, float], target: Dict[str, float]) -> List[Tuple[str, str, float]]:
    """Pair homes that carry too much with homes that carry too little."""
    over = {h: round(actual.get(h, 0.0) - target.get(h, 0.0), 2) for h in set(actual) | set(target)}
    give = sorted([(h, v) for h, v in over.items() if v > TOLERANCE], key=lambda x: -x[1])
    take = sorted([(h, -v) for h, v in over.items() if v < -TOLERANCE], key=lambda x: -x[1])
    out: List[Tuple[str, str, float]] = []
    gi = ti = 0
    give = [list(x) for x in give]
    take = [list(x) for x in take]
    while gi < len(give) and ti < len(take):
        amt = round(min(give[gi][1], take[ti][1]), 2)
        if amt > TOLERANCE:
            out.append((give[gi][0], take[ti][0], amt))
        give[gi][1] = round(give[gi][1] - amt, 2)
        take[ti][1] = round(take[ti][1] - amt, 2)
        if give[gi][1] <= TOLERANCE:
            gi += 1
        if take[ti][1] <= TOLERANCE:
            ti += 1
    return out


def build_costing(
    lines: List[CostLine],
    rules: List[Rule],
    mapping: Dict[str, str],
    master: Dict[str, List[Dict[str, str]]],
    exclude_rows: Optional[Set[int]] = None,
) -> Dict[str, pd.DataFrame]:
    """`exclude_rows`: payroll row numbers already moved by an OTHER-line transfer.
    They are left out of the rules so no line is transferred twice."""
    reverse = {v: k for k, v in mapping.items()}
    assigned, warnings = assign_rules(lines, rules, master)
    rule_by_name = {r.name: r for r in rules}
    exclude = exclude_rows or set()
    skipped = [lines[i] for i in assigned if lines[i].row_no in exclude]
    if skipped:
        names = sorted({assigned[i] for i in assigned if lines[i].row_no in exclude})
        warnings.append({"Rule": ", ".join(names), "Issue": (
            f"{len(skipped)} line(s) totalling \u00a3{sum(x.amount for x in skipped):,.2f} were left out of the rules "
            "because they are already moved by an OTHER-line transfer.")})
    assigned = {i: r for i, r in assigned.items() if lines[i].row_no not in exclude}

    groups: Dict[Tuple[str, str, str, str], List[CostLine]] = defaultdict(list)
    for i, ln in enumerate(lines):
        if i in assigned:
            groups[(assigned[i], ln.period, ln.cost_type, ln.suffix)].append(ln)

    detail, journal = [], []
    for (rname, period, ctype, suffix), items in sorted(groups.items()):
        rule = rule_by_name[rname]
        share = normalise_weights(rule.weights)
        total = round(sum(x.amount for x in items), 2)
        actual: Dict[str, float] = defaultdict(float)
        for x in items:
            actual[x.home] += x.amount
        actual = {h: round(v, 2) for h, v in actual.items()}
        # largest-remainder rounding so targets add up to the pool total to the penny
        target = {h: round(total * s, 2) for h, s in share.items()}
        drift = round(total - sum(target.values()), 2)
        if abs(drift) > 0 and target:
            target[max(share, key=share.get)] = round(target[max(share, key=share.get)] + drift, 2)
        for h in sorted(set(actual) | set(target)):
            detail.append({
                "Rule": rname, "Period": period, "Cost Type": ctype, "Dept Suffix": suffix,
                "Home": h, "Share %": round(share.get(h, 0.0) * 100, 2),
                "Currently Posted": actual.get(h, 0.0), "Should Carry": target.get(h, 0.0),
                "Movement": round(target.get(h, 0.0) - actual.get(h, 0.0), 2),
            })
        for src, dst, amt in _net_transfers(actual, target):
            journal.append({
                "Period": period, "Rule": rname, "Cost Type": ctype, "Dept Suffix": suffix,
                "From Home": src, "Credit Nominal": f"{reverse.get(src, '?')}{suffix}",
                "To Home": dst, "Debit Nominal": f"{reverse.get(dst, '?')}{suffix}",
                "Amount": amt,
            })

    people = []
    for (rname, period, ctype, suffix), items in groups.items():
        for x in items:
            people.append({"Rule": rname, "Staff Name": x.staff_name, "Posted Home": x.home, "Nominal": x.nominal,
                           "Cost Type": ctype, "Amount": x.amount})
    people_df = (pd.DataFrame(people).groupby(["Rule", "Staff Name", "Posted Home", "Nominal", "Cost Type"], as_index=False)["Amount"].sum().round(2)
                 if people else pd.DataFrame(columns=["Rule", "Staff Name", "Posted Home", "Nominal", "Cost Type", "Amount"]))
    unmapped = sorted({ln.prefix for ln in lines if ln.home == "UNKNOWN"})
    for p in unmapped:
        warnings.append({"Rule": "(mapping)", "Issue": f"Nominal prefix {p} is not in the home mapping; lines on it are ignored."})

    jdf = pd.DataFrame(journal, columns=["Period", "Rule", "Cost Type", "Dept Suffix", "From Home", "Credit Nominal", "To Home", "Debit Nominal", "Amount"])
    summary = (jdf.groupby(["Rule", "From Home", "To Home"], as_index=False)["Amount"].sum().round(2)
               if not jdf.empty else pd.DataFrame(columns=["Rule", "From Home", "To Home", "Amount"]))
    return {
        "journal": jdf, "summary": summary, "rule_table": rule_transfer_table(jdf), "detail": pd.DataFrame(detail),
        "people": people_df, "warnings": pd.DataFrame(warnings, columns=["Rule", "Issue"]),
    }


def rule_transfer_table(journal: pd.DataFrame) -> pd.DataFrame:
    """Rule transfers in the same Period / From / To layout, one column per cost type."""
    cols = ["Period", "Rule", "From Home", "To Home", "NIC", "Pension", "Wages"]
    if journal.empty:
        return pd.DataFrame(columns=cols)
    wide = journal.pivot_table(index=["Period", "Rule", "From Home", "To Home"], columns="Cost Type",
                               values="Amount", aggfunc="sum").reset_index()
    wide = wide.rename(columns={"PENSION": "Pension", "WAGES": "Wages"})
    for c in ("NIC", "Pension", "Wages"):
        if c not in wide:
            wide[c] = float("nan")
    wide = (wide.assign(_p=pd.to_numeric(wide["Period"], errors="coerce"))
            .sort_values(["_p", "Rule", "From Home", "To Home"]).drop(columns="_p"))
    return wide[cols].round(2)


# ----------------------------------------------------------------- output --

def write_costing_workbook(
    path: Path,
    result: Dict[str, pd.DataFrame],
    rules: List[Rule],
    no_code: List[Dict[str, Any]],
    mapping: Dict[str, str],
    final_transfers: Optional[pd.DataFrame] = None,
) -> Path:
    """Sheet order: Final_Transfers (regular OTHER-line transfers, ready to share), then the
    new rules kept separate (Rule_Transfers, Journal, Allocation_Detail ...)."""
    from nominal_logic import autosize_and_style

    rule_rows = []
    for r in rules:
        share = normalise_weights(r.weights)
        rule_rows.append({
            "Rule": r.name, "Type": r.kind, "Person": r.person, "Dept Suffixes": ", ".join(r.suffixes),
            "Job Roles": ", ".join(r.job_roles), "Paying Homes": ", ".join(r.paying_homes),
            "Split": ", ".join(f"{h} {s * 100:.2f}%" for h, s in share.items()), "Note": r.note,
        })
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        if final_transfers is not None:
            final_transfers.to_excel(w, sheet_name="Final_Transfers", index=False)
        result["rule_table"].to_excel(w, sheet_name="Rule_Transfers", index=False)
        result["summary"].to_excel(w, sheet_name="Rule_Summary", index=False)
        result["journal"].to_excel(w, sheet_name="Journal", index=False)
        result["detail"].to_excel(w, sheet_name="Allocation_Detail", index=False)
        result["people"].to_excel(w, sheet_name="Staff_Included", index=False)
        pd.DataFrame(rule_rows).to_excel(w, sheet_name="Rules_Used", index=False)
        pd.DataFrame(no_code).to_excel(w, sheet_name="Rows_Without_Nominal", index=False)
        result["warnings"].to_excel(w, sheet_name="Warnings", index=False)
        pd.DataFrame([{"Nominal Prefix": k, "Home": v} for k, v in sorted(mapping.items())]).to_excel(w, sheet_name="Nominal_Mapping", index=False)
    autosize_and_style(path)
    return path
