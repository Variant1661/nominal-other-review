"""Fictional data only."""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from department_costing import CostLine, Rule, build_costing, rule_transfer_table  # noqa: E402
from nominal_logic import DEFAULT_HOME_MAPPING as MAP, final_transfer_table  # noqa: E402


def cl(row, nominal, amount, name="Ann Lee", ctype="WAGES"):
    return CostLine(row, nominal, nominal[0], nominal[1:], MAP[nominal[0]], amount, ctype, "BASIC RATE",
                    name, name.lower(), "7", "x")


KITCHEN = Rule("Kitchen", "pool", {"CRE": 50, "PET": 50}, suffixes=["05"], paying_homes=["CRE", "PET"])


def test_final_transfer_table_layout_and_blanks():
    t = pd.DataFrame([
        {"Period": "10", "Transfer From": "ALI", "Transfer To": "ATW", "Cost Type": "WAGES", "Amount": 5.0},
        {"Period": "7", "Transfer From": "ALI", "Transfer To": "ATW", "Cost Type": "NIC", "Amount": 1.5},
        {"Period": "7", "Transfer From": "ALI", "Transfer To": "ATW", "Cost Type": "WAGES", "Amount": 10.0},
    ])
    out = final_transfer_table(t)
    assert list(out.columns) == ["Period", "Transfer From", "Transfer To", "NIC", "Wages"]
    assert out["Period"].tolist() == ["7", "10"]  # numeric order
    assert out.loc[0, "NIC"] == 1.5 and out.loc[0, "Wages"] == 10.0
    assert pd.isna(out.loc[1, "NIC"])  # blank, not 0
    assert final_transfer_table(pd.DataFrame()).empty


def test_pool_split_nets_to_zero_and_signs_follow_reversals():
    lines = [cl(1, "205", 1000.0), cl(2, "205", -100.0)]
    res = build_costing(lines, [KITCHEN], MAP, {})
    assert res["detail"].groupby("Rule")["Movement"].sum().round(2).iloc[0] == 0
    assert res["summary"]["Amount"].iloc[0] == 450.0  # (1000-100) / 2


def test_lines_already_moved_by_other_transfer_are_excluded_once():
    lines = [cl(1, "205", 1000.0), cl(2, "305", 200.0, name="Bob Roe")]
    kitchen = Rule("Kitchen", "pool", {"CRE": 50, "PET": 50}, suffixes=["05"], paying_homes=["CRE", "PET"])
    with_all = build_costing(lines, [kitchen], MAP, {})
    skipped = build_costing(lines, [kitchen], MAP, {}, exclude_rows={2})
    assert with_all["summary"]["Amount"].sum() == 400.0
    assert skipped["summary"]["Amount"].sum() == 500.0
    assert "already moved" in skipped["warnings"]["Issue"].iloc[0]
    assert "Bob Roe" not in skipped["people"]["Staff Name"].tolist()


def test_person_rule_days_become_percentages():
    lines = [cl(1, "209", 500.0, name="Cy Dee")]
    rule = Rule("Coordinator", "person", {"PET": 2, "CRE": 2, "BUN": 1}, person="Cy Dee")
    res = build_costing(lines, [rule], MAP, {})
    moved = {r["To Home"]: r["Amount"] for _, r in res["journal"].iterrows()}
    assert moved == {"PET": 200.0, "BUN": 100.0}
    assert rule_transfer_table(res["journal"])["Wages"].sum() == 300.0
