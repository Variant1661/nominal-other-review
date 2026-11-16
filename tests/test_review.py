"""Fictional data only."""
import csv
import sys
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from nominal_logic import (  # noqa: E402
    DEFAULT_HOME_MAPPING as MAP,
    apply_missing_nominals,
    build_review_rows,
    infer_staff_base,
    is_other_category,
    load_staff_master,
    missing_nominal_key,
    read_lines,
    staff_lookup_keys,
)
from review_helpers import (  # noqa: E402
    correction_staff_df,
    missing_staff_df,
    save_correction_staff_nominals,
    save_staff_overrides_from_df,
    staff_base_df,
    unusable_base_homes,
)

HEADER = ["Code", "Reference", "Debit", "Credit", "Description"]


def write_csv(tmp_path, rows):
    p = tmp_path / "pay.csv"
    with p.open("w", newline="") as f:
        csv.writer(f).writerows([HEADER, *rows])
    return p


def test_reversal_nets_to_zero_in_transfer_summary(tmp_path):
    p = write_csv(tmp_path, [
        ["204", "r", "500", "", "BASIC for Jane Doe - Period 3"],
        ["304", "r", "100", "", "W/DAY OTHER for Jane Doe - Period 3"],
        ["304", "r", "", "100", "W/DAY OTHER for Jane Doe - Period 3"],
    ])
    lines = read_lines(p)
    base = infer_staff_base(lines, MAP, {}, {})
    review, transfer, _, _ = build_review_rows(lines, base, MAP)
    assert review["Amount"].tolist() == [100.0, -100.0]
    assert transfer["Amount"].sum() == 0


def test_blank_code_is_not_inherited_from_previous_row(tmp_path):
    p = write_csv(tmp_path, [
        ["204", "r", "500", "", "BASIC for Jane Doe - Period 3"],
        ["", "r", "40", "", "W/DAY OTHER for Bob Roe - Period 3"],
    ])
    lines = read_lines(p)
    assert [l.nominal_code for l in lines] == ["204", ""]
    assert not missing_staff_df(lines, {}, MAP).empty


def test_correction_keys_unique_for_identical_lines(tmp_path):
    row = ["", "r", "40", "", "W/DAY OTHER for Bob Roe - Period 3"]
    lines = read_lines(write_csv(tmp_path, [row, row]))
    assert len({missing_nominal_key(l) for l in lines}) == 2
    apply_missing_nominals(lines, {missing_nominal_key(lines[0]): "304"})
    assert [l.nominal_code for l in lines] == ["304", ""]


def test_staff_correction_save_only_touches_edited_staff(tmp_path):
    lines = read_lines(write_csv(tmp_path, [
        ["", "r", "10", "", "W/DAY OTHER for Jane Doe - Period 3"],
        ["", "r", "20", "", "W/DAY OTHER for Jane Doe - Period 3"],
        ["", "r", "30", "", "W/DAY OTHER for Bob Roe - Period 3"],
    ]))
    k = [missing_nominal_key(l) for l in lines]
    corr = {k[0]: "208", k[1]: "504", k[2]: "404"}  # Jane has two different codes
    original = correction_staff_df(lines, corr, MAP)
    edited = original.copy()
    edited.loc[edited["Staff Name"] == "Bob Roe", "Nominal"] = "405"
    assert save_correction_staff_nominals(edited, lines, corr, original) == []
    assert (corr[k[0]], corr[k[1]], corr[k[2]]) == ("208", "504", "405")


def test_untouched_base_rows_do_not_become_overrides(tmp_path):
    lines = read_lines(write_csv(tmp_path, [
        ["204", "r", "500", "", "BASIC for Jane Doe - Period 3"],
        ["304", "r", "10", "", "W/DAY OTHER for Jane Doe - Period 3"],
    ]))
    base = infer_staff_base(lines, MAP, {}, {})
    table = staff_base_df(base, MAP)
    overrides = {}
    assert save_staff_overrides_from_df(table.copy(), overrides, table) == []
    assert overrides == {}
    edited = table.copy()
    edited.loc[0, "Base Nominal"] = "504"
    save_staff_overrides_from_df(edited, overrides, table)
    assert overrides == {"jane doe": "504"}
    cleared = edited.copy()
    cleared.loc[0, "Base Nominal"] = ""
    save_staff_overrides_from_df(cleared, overrides, edited)
    assert overrides == {}


def test_unmapped_base_prefix_blocks(tmp_path):
    lines = read_lines(write_csv(tmp_path, [["304", "r", "10", "", "W/DAY OTHER for Jane Doe - Period 3"]]))
    base = infer_staff_base(lines, MAP, {}, {"jane doe": "804"})
    assert len(unusable_base_homes(staff_base_df(base, MAP))) == 1


def test_master_exact_name_beats_weaker_variants_and_ambiguity_is_flagged(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["Employee Code", "Home", "Full Name"])
    ws.append([1, "Crescent", "John Smith"])
    ws.append([2, "Peter's Place", "John Paul Smith"])  # weak key "john smith" collides
    ws.append([3, "Bunyan Lodge", "Ann Lee"])
    ws.append([4, "Atwell House", "Ann Lee"])  # identical name, different person
    path = tmp_path / "m.xlsx"
    wb.save(path)
    master = load_staff_master(path, MAP)
    assert master["john smith"]["employee_code"] == "1"
    assert not master["john smith"]["ambiguous"]
    assert master["ann lee"]["ambiguous"]
    keys = staff_lookup_keys("John Paul Smith")
    assert keys == staff_lookup_keys("John Paul Smith") and keys[0] == "john paul smith"


def test_other_is_a_whole_word():
    assert is_other_category("W/DAY OTHER") and is_other_category("OTHERS")
    assert not is_other_category("MOTHERS DAY PAY")


def test_unreadable_amount_is_reported(tmp_path):
    p = write_csv(tmp_path, [["204", "r", "1,234.56 CR", "", "BASIC for Jane Doe - Period 3"]])
    with pytest.raises(ValueError, match="row 2"):
        read_lines(p)


def test_periods_sort_numerically(tmp_path):
    lines = read_lines(write_csv(tmp_path, [
        ["204", "r", "50", "", "W/DAY OTHER for Jane Doe - Period 10"],
        ["204", "r", "40", "", "W/DAY OTHER for Jane Doe - Period 2"],
        ["304", "r", "5", "", "BASIC for Jane Doe - Period 2"],
    ]))
    base = infer_staff_base(lines, MAP, {}, {"jane doe": "304"})
    _, transfer, _, _ = build_review_rows(lines, base, MAP)
    assert transfer["Period"].tolist() == ["2", "10"]
