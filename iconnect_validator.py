#!/usr/bin/env python3
"""
iConnect Payroll Extract Validator — Enfield Pension Fund 2026/27
Validates one or more CSV payroll extracts against the File Completion Guide
and writes all flags to an Excel workbook.

Within-file checks run on every file independently (mandatory/conditional
fields, formats, contribution bands, 50/50 logic, etc). When more than one
file is supplied, an additional cross-period pass chains the periods
together per member (matched on NI_NUMBER + PAY_REF_1-3) to catch things a
single file can never reveal on its own: cumulative figures that don't
add up between periods, members who vanish without a DATE_OF_LEAVING,
opted-out members who keep accruing pay, and so on.

Usage:
    python iconnect_validator.py <extract.csv> [-o output.xlsx]
    python iconnect_validator.py april.csv may.csv june.csv [-o output.xlsx]
    python iconnect_validator.py --dir ./extracts/ [-o output.xlsx]

Files are sorted by PAYROLL_PERIOD_END_DATE automatically — pass them in
any order or filename convention.
"""

import argparse
import csv
import glob
import re
import sys
import os
from datetime import datetime
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("ERROR: openpyxl is required — pip install openpyxl")

# ── Required column order (63 columns) ────────────────────────────────────────────
EXPECTED_COLUMNS = [
    "NI_NUMBER", "PAY_REF_1", "PAY_REF_2", "PAY_REF_3",
    "ADD_LINE_1", "ADD_LINE_2", "ADD_LINE_3", "ADD_LINE_4", "ADD_LINE_5",
    "POSTCODE", "EMAIL_ADDRESS", "TELEPHONE_NUMBER", "MOBILE_NUMBER",
    "WORKS_PLACE_NAME", "WORKS_ADD_LINE_1", "WORKS_ADD_LINE_2",
    "WORKS_ADD_LINE_3", "WORKS_ADD_LINE_4", "WORKS_ADD_LINE_5",
    "WORKS_POSTCODE", "WORKS_EMAIL_ADDRESS",
    "DATE_OF_LEAVING", "PAYROLL_PERIOD_END_DATE",
    "ADDITIONAL_CONTRIBUTIONS_1", "ADDITIONAL_CONTRIBUTIONS_2",
    "EMPLOYMENT_BREAK_START", "EMPLOYMENT_BREAK_END",
    "FILLER_1",
    "EMPLOYMENT_BREAK_REASON",
    "SURNAME", "FORENAMES", "GENDER", "DOB", "MARITAL_STATUS", "TITLE",
    "FILLER_2",
    "TAXABLE_EARNINGS", "ANNUAL_PENSIONABLE_SALARY", "PENSIONABLE_PAY",
    "EFFECTIVE_DATE", "DATE_JOINED_PENSION_SCHEME", "JOB_TITLE",
    "PART_TIME_HOURS_EFFECTIVE_DATE", "PART_TIME_HOURS", "PART_TIME_INDICATOR",
    "WHOLE_TIME_EQUIVALENT_HOURS",
    "EMPLOYEES_MAIN_SECTION_CONTS", "EMPLOYERS_CONTS", "SCHEME_CONT_RATE",
    "OPT_OUT_DATE", "OPT_IN_DATE",
    "MAIN_SECTION_CUMULATIVE_PEN_PAY",
    "5050_SECTION_CUMULATIVE_PEN_PAY",
    "FTE_FINAL_PAY",
    "CUMULATIVE_EMPLOYEES_MAIN_SECTION_SCHEME_CONTS",
    "CUMULATIVE_EMPLOYERS_SCHEME_CONTS",
    "REASON_FOR_LEAVING", "CUMULATIVE_SCAPCs", "CUMULATIVE_APCs",
    "EMPLOYEES_5050_CONTS", "CUMULATIVE_EMPLOYEES_5050_CONTS",
    "SCAPCs", "APCs",
]

# ── 2026/27 contribution bands (low, high, main%, 50/50%) ───────────────────────
CONTRIBUTION_BANDS = [
    (0,         18400,              5.50, 2.75),
    (18400.01,  29000,              5.80, 2.90),
    (29000.01,  47300,              6.50, 3.25),
    (47300.01,  59800,              6.80, 3.40),
    (59800.01,  84000,              8.50, 4.25),
    (84000.01,  119100,             9.90, 4.95),
    (119100.01, 140400,            10.50, 5.25),
    (140400.01, 210700,            11.40, 5.70),
    (210700.01, float("inf"),      12.50, 6.25),
]

VALID_BREAK_REASONS     = {"A", "E", "M", "S", "U", "Y"}
VALID_MARITAL_STATUS    = {"S", "M", "C", "D", "W"}
VALID_PART_TIME_IND     = {"Y", "C"}
VALID_GENDER            = {"M", "F"}

# NI format: 2 letters, 6 digits, 1 letter (simplified but accurate)
NI_RE       = re.compile(r"^[A-CEGHJ-PR-TW-Z][A-CEGHJ-NPR-TW-Z]\d{6}[A-D]$", re.I)
POSTCODE_RE = re.compile(r"^[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}$", re.I)
DATE_RE     = re.compile(r"^\d{2}/\d{2}/\d{4}$")

# Cumulative field ↔ this-period field pairs used for cross-period checks.
# section: None = always checked, "main" = only when not in 50/50,
# "5050" = only when in 50/50.
CUMULATIVE_PAIRS = [
    ("MAIN_SECTION_CUMULATIVE_PEN_PAY",                  "PENSIONABLE_PAY",              "main"),
    ("5050_SECTION_CUMULATIVE_PEN_PAY",                  "PENSIONABLE_PAY",              "5050"),
    ("CUMULATIVE_EMPLOYEES_MAIN_SECTION_SCHEME_CONTS",   "EMPLOYEES_MAIN_SECTION_CONTS", None),
    ("CUMULATIVE_EMPLOYERS_SCHEME_CONTS",                "EMPLOYERS_CONTS",              None),
    ("CUMULATIVE_EMPLOYEES_5050_CONTS",                  "EMPLOYEES_5050_CONTS",         None),
    ("CUMULATIVE_APCs",                                  "APCs",                         None),
    ("CUMULATIVE_SCAPCs",                                "SCAPCs",                       None),
]

MAX_GAP_DAYS = 40   # slack above a calendar month for "consecutive" periods

# ── Flag catalogue ───────────────────────────────────────────────────────────────
# code → (severity_category, short_description)
FLAGS = {
    "F001": ("MANDATORY_BLANK",         "Mandatory field is blank"),
    "F002": ("NI_FORMAT",               "NI_NUMBER does not match format AA123456A"),
    "F003": ("DATE_FORMAT",             "Date field not in DD/MM/YYYY format"),
    "F004": ("DATE_INVALID",            "Date is formatted correctly but is not a valid calendar date"),
    "F005": ("POSTCODE_FORMAT",         "POSTCODE does not appear to be a valid UK postcode"),
    "F006": ("GENDER_INVALID",          "GENDER must be M or F"),
    "F007": ("NUMERIC_FORMAT",          "Numeric field contains commas, £ signs or non-numeric characters"),
    "F008": ("PERIOD_DATE_MISMATCH",    "PAYROLL_PERIOD_END_DATE differs across rows — must be identical on every row"),
    "F009": ("FILLER_POPULATED",        "FILLER_1 or FILLER_2 is not blank — must always be empty"),
    "F010": ("COLUMN_COUNT",            "File does not have exactly 63 columns"),
    "F011": ("COLUMN_ORDER",            "Column headers are not in the required sequence"),
    "F012": ("MISSING_CONDITIONAL",     "Conditional field is required but missing (e.g. EFFECTIVE_DATE when salary is given)"),
    "F013": ("LEAVER_NO_REASON",        "DATE_OF_LEAVING is set but REASON_FOR_LEAVING is blank"),
    "F014": ("BREAK_NO_REASON",         "EMPLOYMENT_BREAK_START is set but EMPLOYMENT_BREAK_REASON is blank"),
    "F015": ("BREAK_REASON_INVALID",    "EMPLOYMENT_BREAK_REASON is not a valid code (A/E/M/S/U/Y)"),
    "F016": ("MARITAL_STATUS_INVALID",  "MARITAL_STATUS must be S, M, C, D or W"),
    "F017": ("PT_INDICATOR_INVALID",    "PART_TIME_INDICATOR must be Y or C"),
    "F018": ("PT_FIELDS_INCONSISTENT",  "Part-time indicator set but PART_TIME_HOURS or WTE_HOURS is missing"),
    "F019": ("OPT_OUT_AND_IN_SAME_ROW", "OPT_OUT_DATE and OPT_IN_DATE both populated on the same row"),
    "F020": ("OPT_OUT_RATE_NOT_ZERO",   "SCHEME_CONT_RATE should be 0 after opt-out"),
    "F021": ("TAXABLE_EARNINGS_CHECK",  "TAXABLE_EARNINGS ≠ PENSIONABLE_PAY minus employee contributions (±0.02 tolerance)"),
    "F022": ("CONT_RATE_BAND_CHECK",    "SCHEME_CONT_RATE does not match the 2026/27 band for ANNUAL_PENSIONABLE_SALARY"),
    "F023": ("5050_MAIN_OVERLAP",       "50/50 member: EMPLOYEES_MAIN_SECTION_CONTS should be 0"),
    "F024": ("5050_PAY_WRONG_COL",      "50/50 member: pay should accumulate in 5050_SECTION_CUMULATIVE_PEN_PAY, not main column"),
    "F025": ("CUMULATIVE_BELOW_PERIOD", "A cumulative year-to-date value is less than the current period value"),
    "F026": ("DOL_USED_FOR_OPTOUT",     "DATE_OF_LEAVING set on opt-out row — use OPT_OUT_DATE for opt-outs"),
    "F027": ("SURNAME_NOT_CAPS",        "SURNAME must be in CAPITAL LETTERS"),
    "F028": ("FORENAMES_NOT_CAPS",      "FORENAMES must be in CAPITAL LETTERS"),
    "F029": ("TITLE_NOT_CAPS",          "TITLE must be in CAPITAL LETTERS"),
    "F030": ("JOB_TITLE_NOT_CAPS",      "JOB_TITLE must be in CAPITAL LETTERS"),
    "F031": ("UNEXPECTED_NEGATIVE",     "Negative value found in a field where negatives are not expected"),
    "F032": ("PAY_REF_TOO_LONG",        "PAY_REF exceeds maximum 12 characters"),
    "F033": ("MISSING_COLUMN",          "One or more required columns are absent from the header"),
    "F034": ("DOL_AND_OPTOUT_SAME_ROW", "DATE_OF_LEAVING and OPT_OUT_DATE both set — only populate OPT_OUT_DATE for opt-outs"),
    # Cross-period checks — only raised when 2+ files are supplied
    "F035": ("MISSING_PERIOD_GAP",       "Period end dates are not consecutive months — a monthly submission may be missing"),
    "F036": ("MEMBER_VANISHED",          "Member present in the previous period is missing here with no DATE_OF_LEAVING"),
    "F037": ("NEW_MEMBER_NO_JOIN_DATE",  "Member appears mid-year without DATE_JOINED_PENSION_SCHEME or OPT_IN_DATE"),
    "F038": ("CUMULATIVE_MISMATCH",      "Cumulative ≠ previous period's cumulative + this period's figure (±0.02)"),
    "F039": ("CUMULATIVE_DECREASED",     "Cumulative fell between periods outside an opt-out refund row"),
    "F040": ("POST_OPTOUT_ACTIVITY",     "New pay/contributions or changed cumulatives after a reported opt-out"),
    "F041": ("POST_OPTOUT_RATE_NOT_ZERO","SCHEME_CONT_RATE is not 0 in a period after the opt-out"),
    "F042": ("OPT_IN_CARRIED_FORWARD",   "OPT_IN_DATE still populated in a period after the one it was first reported in"),
    "F043": ("LEAVER_REAPPEARED",        "Member with a prior DATE_OF_LEAVING reappears with no new join/opt-in date"),
    "F044": ("STATIC_DATA_CHANGED",      "DOB or GENDER changed for the same member between periods"),
}

FIX_GUIDANCE = {
    "F001": "Ensure every mandatory field is populated on every row. For numeric fields with no value to report, use 0 (not blank).",
    "F002": "NI number must be 2 letters, 6 digits, 1 letter — e.g. AB123456C — with no spaces. Check source data.",
    "F003": "All date fields must use DD/MM/YYYY format with leading zeros — e.g. 01/04/2026.",
    "F004": "The date format is correct but the date is impossible (e.g. 30/02/2026). Check for transposed day/month.",
    "F005": "Postcode must follow UK format — e.g. EN1 3XY. Check for missing spaces, incorrect letters or typos.",
    "F006": "GENDER accepts M or F only. If the system holds another value, contact the Fund for the accepted reporting code.",
    "F007": "Remove all commas (,), £ signs, spaces and % symbols from numeric fields. Use plain decimals — e.g. 2500.00.",
    "F008": "PAYROLL_PERIOD_END_DATE must be identical on every row. A file with mixed dates will be rejected by i-Connect.",
    "F009": "FILLER_1 (col 28) and FILLER_2 (col 36) must always be blank. Remove any values; keep the column header.",
    "F010": "Add or remove columns until exactly 63 are present. Use the Fund-issued template as the basis.",
    "F011": "Restore the column order to match the template exactly. Do not move or rename headers.",
    "F012": "When ANNUAL_PENSIONABLE_SALARY is provided, EFFECTIVE_DATE is mandatory. Also: new starters need DATE_JOINED_PENSION_SCHEME.",
    "F013": "Provide a REASON_FOR_LEAVING value (e.g. Resignation, Retirement, Redundancy) whenever DATE_OF_LEAVING is set.",
    "F014": "Provide EMPLOYMENT_BREAK_REASON whenever EMPLOYMENT_BREAK_START is populated.",
    "F015": "EMPLOYMENT_BREAK_REASON must be one of: A (Leave of Absence), E (Education), M (Parental), S (Strike), U (Unauthorised), Y (Maternity/Paternity).",
    "F016": "MARITAL_STATUS must be: S (Single), M (Married), C (Civil Partnership), D (Divorced), W (Widowed). Leave blank if unknown.",
    "F017": "PART_TIME_INDICATOR accepts Y (part-time) or C (casual) only. Leave blank for whole-time staff.",
    "F018": "When PART_TIME_INDICATOR = Y, also provide PART_TIME_HOURS and WHOLE_TIME_EQUIVALENT_HOURS.",
    "F019": "OPT_OUT_DATE and OPT_IN_DATE cannot both appear on the same row. Use separate periods for each event.",
    "F020": "Once a member has opted out, set SCHEME_CONT_RATE to 0 for that row and subsequent rows until they leave.",
    "F021": "TAXABLE_EARNINGS = PENSIONABLE_PAY minus employee contributions (LGPS contributions reduce taxable pay). Recalculate.",
    "F022": "Use the member's actual annual pensionable pay against the 2026/27 band table (see 'Contribution Rates' sheet) to set SCHEME_CONT_RATE. Update to 2026/27 bands from 1 April 2026.",
    "F023": "50/50 members: employee contributions go in EMPLOYEES_5050_CONTS. EMPLOYEES_MAIN_SECTION_CONTS must be 0.",
    "F024": "50/50 members: pensionable pay accumulates in 5050_SECTION_CUMULATIVE_PEN_PAY; MAIN_SECTION_CUMULATIVE_PEN_PAY should be 0.",
    "F025": "Cumulative year-to-date figures must be >= the current-period figures. Check for data-entry or reset errors.",
    "F026": "For opt-outs, use OPT_OUT_DATE. Do not populate DATE_OF_LEAVING unless the member has also left employment in the same period.",
    "F027": "SURNAME must be supplied in CAPITAL LETTERS as required by the i-Connect specification.",
    "F028": "FORENAMES must be supplied in CAPITAL LETTERS.",
    "F029": "TITLE must be in CAPITAL LETTERS — e.g. MR, MRS, MS, DR.",
    "F030": "JOB_TITLE must be in CAPITAL LETTERS.",
    "F031": "Negative values are only permitted in opt-out refund rows (OPT_OUT_DATE is populated and the member is within the 3-month refund window). Check whether this row qualifies.",
    "F032": "PAY_REF values must be at most 12 characters. Agree the reference format with the Fund and keep it stable.",
    "F033": "Restore all missing column headers from the Fund-issued template. All 63 columns must be present.",
    "F034": "Use OPT_OUT_DATE for opt-outs. Only use DATE_OF_LEAVING when the member has physically left employment.",
    "F035": "Check whether a monthly submission is missing between these two period end dates. If the gap is intentional (e.g. an annually-paid employer), confirm with the Fund.",
    "F036": "Confirm whether this member left, transferred, or was omitted in error. If they left, resubmit with DATE_OF_LEAVING and REASON_FOR_LEAVING populated — never simply drop a member from the extract.",
    "F037": "New starters and re-joiners must have DATE_JOINED_PENSION_SCHEME or OPT_IN_DATE populated in the period they first appear.",
    "F038": "Cumulative year-to-date figures should equal the previous period's cumulative plus this period's amount. Check for a missed period, a miscalculation, or a manual override.",
    "F039": "A cumulative figure has fallen between periods. This is only expected on an opt-out refund row (PENSIONABLE_PAY=0, contributions reversed). Otherwise investigate a data-entry or reset error.",
    "F040": "Opted-out members must not accrue new pensionable pay, contributions, or cumulative changes after the opt-out is reported. Confirm the member remains opted out or process a rejoin via OPT_IN_DATE.",
    "F041": "SCHEME_CONT_RATE must be 0 in every period after an opt-out, until the member rejoins or leaves.",
    "F042": "OPT_IN_DATE should only appear in the pay period the opt-in/re-enrolment is first reported. Clear it in subsequent periods.",
    "F043": "A member with a prior DATE_OF_LEAVING has reappeared. Populate DATE_JOINED_PENSION_SCHEME (new post) or OPT_IN_DATE (rejoin) to explain the reappearance.",
    "F044": "DOB and GENDER should not change for the same member. Verify this is not a data-entry error or a mismatched NI/PAY_REF combination.",
}


# ── Helpers ────────────────────────────────────────────────────────────────────────

def _v(row, col):
    """Get value from row dict, returning '' if missing."""
    return row.get(col, "") or ""


def _blank(val):
    return str(val).strip() == ""


def _to_float(val):
    if val is None or str(val).strip() == "":
        return None
    s = str(val).strip().replace(",", "").replace("£", "").replace("%", "").replace(" ", "")
    try:
        return float(s)
    except ValueError:
        return None


def _numeric_ok(val):
    """True if blank or clean numeric (no commas/£/%)."""
    if _blank(val):
        return True
    s = str(val).strip()
    if "," in s or "£" in s or "%" in s:
        return False
    try:
        float(s)
        return True
    except ValueError:
        return False


def _date_check(val):
    """
    Returns:
        None  — blank (not an error)
        True  — valid DD/MM/YYYY date
        'fmt' — wrong format
        'inv' — right format but impossible date
    """
    s = str(val).strip()
    if not s:
        return None
    if not DATE_RE.match(s):
        return "fmt"
    try:
        datetime.strptime(s, "%d/%m/%Y")
        return True
    except ValueError:
        return "inv"


def _expected_rate(annual_pay, is_50):
    for lo, hi, main, half in CONTRIBUTION_BANDS:
        if lo <= annual_pay <= hi:
            return half if is_50 else main
    return None


def _is_5050(row):
    """Heuristic: member is in 50/50 if 5050 contribution or cumulative pay is non-zero,
    or if SCHEME_CONT_RATE matches a 50/50 band rate."""
    e50 = _to_float(_v(row, "EMPLOYEES_5050_CONTS"))
    c50 = _to_float(_v(row, "5050_SECTION_CUMULATIVE_PEN_PAY"))
    if e50 and e50 != 0:
        return True
    if c50 and c50 != 0:
        return True
    rate = _to_float(_v(row, "SCHEME_CONT_RATE"))
    if rate:
        for _, _, _, half in CONTRIBUTION_BANDS:
            if abs(rate - half) < 0.02:
                return True
    return False


def _member_key(row):
    """Matching key per the guide: NI_NUMBER + PAY_REF_1-3 combination."""
    return (
        str(_v(row, "NI_NUMBER")).strip().upper(),
        str(_v(row, "PAY_REF_1")).strip().upper(),
        str(_v(row, "PAY_REF_2")).strip().upper(),
        str(_v(row, "PAY_REF_3")).strip().upper(),
    )


def _first_valid_period_date(rows):
    """First row-order PAYROLL_PERIOD_END_DATE that parses as a valid date."""
    for row in rows:
        v = _v(row, "PAYROLL_PERIOD_END_DATE")
        if _date_check(v) is True:
            return datetime.strptime(str(v).strip(), "%d/%m/%Y")
    return None


# ── Single-file validation engine ───────────────────────────────────────────────────

def validate(csv_path):
    """Return (file_flags, row_flags, rows, actual_headers) for one CSV file."""

    file_flags = []   # file-level issues
    row_flags  = []   # per-row issues

    with open(csv_path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        headers = list(reader.fieldnames or [])
        rows    = list(reader)

    # ── File-level checks ─────────────────────────────────────────────────────────────────

    if len(headers) != 63:
        file_flags.append(dict(row="FILE", ni="—", field="HEADER",
            code="F010", detail=f"Has {len(headers)} columns; expected 63."))

    missing = [c for c in EXPECTED_COLUMNS if c not in headers]
    if missing:
        file_flags.append(dict(row="FILE", ni="—", field="HEADER",
            code="F033", detail=f"Missing: {', '.join(missing)}"))

    present_in_order  = [c for c in headers   if c in set(EXPECTED_COLUMNS)]
    expected_in_order = [c for c in EXPECTED_COLUMNS if c in set(headers)]
    if present_in_order != expected_in_order:
        file_flags.append(dict(row="FILE", ni="—", field="HEADER",
            code="F011", detail="Column order does not match the required sequence."))

    # Collect all period-end dates for cross-row check
    period_dates = defaultdict(list)

    def flag(rnum, ni, field, code, detail=""):
        row_flags.append(dict(row=rnum, ni=ni, field=field,
                              code=code, detail=detail))

    for idx, row in enumerate(rows, start=2):   # row 1 = header
        ni   = _v(row, "NI_NUMBER")
        rn   = idx

        # NI_NUMBER
        ni_s = str(ni).strip()
        if _blank(ni):
            flag(rn, ni, "NI_NUMBER",    "F001", "NI_NUMBER is blank")
        elif not NI_RE.match(ni_s):
            flag(rn, ni, "NI_NUMBER",    "F002", f"Value: {ni}")

        # PAY_REF_1
        pr1 = _v(row, "PAY_REF_1")
        if _blank(pr1):
            flag(rn, ni, "PAY_REF_1",   "F001", "PAY_REF_1 is blank")
        elif len(str(pr1).strip()) > 12:
            flag(rn, ni, "PAY_REF_1",   "F032", f"Length {len(str(pr1).strip())} > 12")

        for ref in ("PAY_REF_2", "PAY_REF_3"):
            rv = _v(row, ref)
            if not _blank(rv) and len(str(rv).strip()) > 12:
                flag(rn, ni, ref,        "F032", f"Length {len(str(rv).strip())} > 12")

        # Mandatory address fields
        for fld in ("ADD_LINE_1", "ADD_LINE_2"):
            if _blank(_v(row, fld)):
                flag(rn, ni, fld,        "F001", f"{fld} is blank")

        # POSTCODE
        pc = _v(row, "POSTCODE")
        if _blank(pc):
            flag(rn, ni, "POSTCODE",     "F001", "POSTCODE is blank")
        elif not POSTCODE_RE.match(str(pc).strip()):
            flag(rn, ni, "POSTCODE",     "F005", f"Value: {pc}")

        # PAYROLL_PERIOD_END_DATE
        pped = _v(row, "PAYROLL_PERIOD_END_DATE")
        if _blank(pped):
            flag(rn, ni, "PAYROLL_PERIOD_END_DATE", "F001", "Blank")
        else:
            dc = _date_check(pped)
            if dc == "fmt":
                flag(rn, ni, "PAYROLL_PERIOD_END_DATE", "F003", f"Value: {pped}")
            elif dc == "inv":
                flag(rn, ni, "PAYROLL_PERIOD_END_DATE", "F004", f"Value: {pped}")
            else:
                period_dates[str(pped).strip()].append(rn)

        # SURNAME / FORENAMES / TITLE / JOB_TITLE — mandatory + CAPS
        for fld, caps_code in (
            ("SURNAME",   "F027"), ("FORENAMES", "F028"),
            ("TITLE",     "F029"), ("JOB_TITLE", "F030"),
        ):
            val = _v(row, fld)
            if _blank(val):
                flag(rn, ni, fld, "F001", f"{fld} is blank")
            elif str(val).strip() != str(val).strip().upper():
                flag(rn, ni, fld, caps_code, f"Value: {val}")

        # GENDER
        g = str(_v(row, "GENDER")).strip().upper()
        if _blank(g):
            flag(rn, ni, "GENDER", "F001", "GENDER is blank")
        elif g not in VALID_GENDER:
            flag(rn, ni, "GENDER", "F006", f"Value: {g}")

        # DOB
        dob = _v(row, "DOB")
        if _blank(dob):
            flag(rn, ni, "DOB", "F001", "DOB is blank")
        else:
            dc = _date_check(dob)
            if dc == "fmt":
                flag(rn, ni, "DOB", "F003", f"Value: {dob}")
            elif dc == "inv":
                flag(rn, ni, "DOB", "F004", f"Value: {dob}")

        # MARITAL_STATUS (optional — validate if present)
        ms = str(_v(row, "MARITAL_STATUS")).strip().upper()
        if ms and ms not in VALID_MARITAL_STATUS:
            flag(rn, ni, "MARITAL_STATUS", "F016", f"Value: {ms}")

        # Mandatory numeric fields
        for fld in (
            "TAXABLE_EARNINGS", "PENSIONABLE_PAY",
            "EMPLOYEES_MAIN_SECTION_CONTS", "EMPLOYERS_CONTS", "SCHEME_CONT_RATE",
            "MAIN_SECTION_CUMULATIVE_PEN_PAY",
            "CUMULATIVE_EMPLOYEES_MAIN_SECTION_SCHEME_CONTS",
            "CUMULATIVE_EMPLOYERS_SCHEME_CONTS",
        ):
            val = _v(row, fld)
            if _blank(val):
                flag(rn, ni, fld, "F001", f"{fld} is blank — use 0 if none")
            elif not _numeric_ok(val):
                flag(rn, ni, fld, "F007", f"Value: {val}")

        # Optional date fields — format-check if present
        for fld in (
            "DATE_OF_LEAVING", "EMPLOYMENT_BREAK_START", "EMPLOYMENT_BREAK_END",
            "EFFECTIVE_DATE", "DATE_JOINED_PENSION_SCHEME",
            "PART_TIME_HOURS_EFFECTIVE_DATE", "OPT_OUT_DATE", "OPT_IN_DATE",
        ):
            val = _v(row, fld)
            if not _blank(val):
                dc = _date_check(val)
                if dc == "fmt":
                    flag(rn, ni, fld, "F003", f"Value: {val}")
                elif dc == "inv":
                    flag(rn, ni, fld, "F004", f"Value: {val}")

        # FILLER columns must be blank
        for fld in ("FILLER_1", "FILLER_2"):
            if not _blank(_v(row, fld)):
                flag(rn, ni, fld, "F009", f"Value: {_v(row, fld)}")

        # Conditional: DATE_OF_LEAVING → REASON_FOR_LEAVING
        dol    = _v(row, "DATE_OF_LEAVING")
        rfl    = _v(row, "REASON_FOR_LEAVING")
        optout = _v(row, "OPT_OUT_DATE")
        optin  = _v(row, "OPT_IN_DATE")

        if not _blank(dol) and _blank(rfl):
            flag(rn, ni, "REASON_FOR_LEAVING", "F013")

        # OPT_OUT_DATE and OPT_IN_DATE on same row
        if not _blank(optout) and not _blank(optin):
            flag(rn, ni, "OPT_OUT_DATE / OPT_IN_DATE", "F019")

        # DATE_OF_LEAVING and OPT_OUT_DATE on same row
        if not _blank(dol) and not _blank(optout):
            flag(rn, ni, "DATE_OF_LEAVING / OPT_OUT_DATE", "F034")

        # SCHEME_CONT_RATE should be 0 after opt-out
        if not _blank(optout):
            rate = _to_float(_v(row, "SCHEME_CONT_RATE"))
            if rate is not None and rate != 0:
                pen = _to_float(_v(row, "PENSIONABLE_PAY")) or 0
                if pen == 0:
                    flag(rn, ni, "SCHEME_CONT_RATE", "F020",
                         f"OPT_OUT_DATE set and PENSIONABLE_PAY=0; SCHEME_CONT_RATE should be 0, got {rate}")

        # ANNUAL_PENSIONABLE_SALARY → EFFECTIVE_DATE mandatory
        aps_raw = _v(row, "ANNUAL_PENSIONABLE_SALARY")
        eff     = _v(row, "EFFECTIVE_DATE")
        if not _blank(aps_raw):
            if not _numeric_ok(aps_raw):
                flag(rn, ni, "ANNUAL_PENSIONABLE_SALARY", "F007", f"Value: {aps_raw}")
            if _blank(eff):
                flag(rn, ni, "EFFECTIVE_DATE", "F012",
                     "EFFECTIVE_DATE required when ANNUAL_PENSIONABLE_SALARY is provided")

        # Part-time fields
        pti      = str(_v(row, "PART_TIME_INDICATOR")).strip().upper()
        pt_hrs   = _v(row, "PART_TIME_HOURS")
        wte_hrs  = _v(row, "WHOLE_TIME_EQUIVALENT_HOURS")

        if pti and pti not in VALID_PART_TIME_IND:
            flag(rn, ni, "PART_TIME_INDICATOR", "F017", f"Value: {pti}")
        if pti == "Y":
            if _blank(pt_hrs):
                flag(rn, ni, "PART_TIME_HOURS", "F018",
                     "PART_TIME_HOURS required when PART_TIME_INDICATOR=Y")
            if _blank(wte_hrs):
                flag(rn, ni, "WHOLE_TIME_EQUIVALENT_HOURS", "F018",
                     "WHOLE_TIME_EQUIVALENT_HOURS required when PART_TIME_INDICATOR=Y")

        for fld in ("PART_TIME_HOURS", "WHOLE_TIME_EQUIVALENT_HOURS"):
            val = _v(row, fld)
            if not _blank(val) and not _numeric_ok(val):
                flag(rn, ni, fld, "F007", f"Value: {val}")

        # Employment break reason
        eb_start  = _v(row, "EMPLOYMENT_BREAK_START")
        eb_reason = str(_v(row, "EMPLOYMENT_BREAK_REASON")).strip().upper()
        if not _blank(eb_start):
            if _blank(eb_reason):
                flag(rn, ni, "EMPLOYMENT_BREAK_REASON", "F014")
            elif eb_reason not in VALID_BREAK_REASONS:
                flag(rn, ni, "EMPLOYMENT_BREAK_REASON", "F015", f"Value: {eb_reason}")

        # 50/50 section checks
        is_50        = _is_5050(row)
        emp_main     = _to_float(_v(row, "EMPLOYEES_MAIN_SECTION_CONTS")) or 0
        emp_50       = _to_float(_v(row, "EMPLOYEES_5050_CONTS"))          or 0
        main_cum_pay = _to_float(_v(row, "MAIN_SECTION_CUMULATIVE_PEN_PAY")) or 0
        cum_50_pay   = _to_float(_v(row, "5050_SECTION_CUMULATIVE_PEN_PAY")) or 0

        if is_50:
            if emp_main != 0:
                flag(rn, ni, "EMPLOYEES_MAIN_SECTION_CONTS", "F023",
                     f"50/50 member: should be 0, got {emp_main}")
            if main_cum_pay != 0:
                flag(rn, ni, "MAIN_SECTION_CUMULATIVE_PEN_PAY", "F024",
                     f"50/50 member: should be 0, got {main_cum_pay}")

        # TAXABLE_EARNINGS cross-check
        pen_pay  = _to_float(_v(row, "PENSIONABLE_PAY"))
        tax_earn = _to_float(_v(row, "TAXABLE_EARNINGS"))
        if pen_pay is not None and tax_earn is not None:
            ee_conts    = emp_50 if is_50 else emp_main
            expected_te = pen_pay - ee_conts
            if abs(tax_earn - expected_te) > 0.02:
                flag(rn, ni, "TAXABLE_EARNINGS", "F021",
                     f"Expected {expected_te:.2f} (pen_pay {pen_pay:.2f} − ee_conts {ee_conts:.2f}), got {tax_earn:.2f}")

        # Contribution rate band check (only when annual salary is known and valid)
        aps_val     = _to_float(aps_raw)
        scheme_rate = _to_float(_v(row, "SCHEME_CONT_RATE"))
        if aps_val is not None and scheme_rate is not None:
            exp_rate = _expected_rate(aps_val, is_50)
            if exp_rate is not None and abs(scheme_rate - exp_rate) > 0.02:
                flag(rn, ni, "SCHEME_CONT_RATE", "F022",
                     f"Annual pay £{aps_val:,.2f} → expected {exp_rate}% ({'50/50' if is_50 else 'main'}), got {scheme_rate}%")

        # Cumulative >= period
        cum_ee  = _to_float(_v(row, "CUMULATIVE_EMPLOYEES_MAIN_SECTION_SCHEME_CONTS")) or 0
        cum_er  = _to_float(_v(row, "CUMULATIVE_EMPLOYERS_SCHEME_CONTS"))              or 0
        er_conts = _to_float(_v(row, "EMPLOYERS_CONTS")) or 0

        if cum_ee < emp_main and emp_main > 0:
            flag(rn, ni, "CUMULATIVE_EMPLOYEES_MAIN_SECTION_SCHEME_CONTS", "F025",
                 f"Cumulative {cum_ee:.2f} < period {emp_main:.2f}")
        if cum_er < er_conts and er_conts > 0:
            flag(rn, ni, "CUMULATIVE_EMPLOYERS_SCHEME_CONTS", "F025",
                 f"Cumulative {cum_er:.2f} < period {er_conts:.2f}")
        if not is_50 and pen_pay and pen_pay > 0 and main_cum_pay < pen_pay:
            flag(rn, ni, "MAIN_SECTION_CUMULATIVE_PEN_PAY", "F025",
                 f"Cumulative {main_cum_pay:.2f} < period pensionable pay {pen_pay:.2f}")
        if is_50 and pen_pay and pen_pay > 0 and cum_50_pay < pen_pay:
            flag(rn, ni, "5050_SECTION_CUMULATIVE_PEN_PAY", "F025",
                 f"Cumulative {cum_50_pay:.2f} < period pensionable pay {pen_pay:.2f}")

        # Unexpected negatives (only allowed in opt-out refund rows)
        is_refund = not _blank(optout)
        if not is_refund:
            for fld in (
                "EMPLOYEES_MAIN_SECTION_CONTS", "EMPLOYERS_CONTS",
                "PENSIONABLE_PAY", "TAXABLE_EARNINGS",
                "MAIN_SECTION_CUMULATIVE_PEN_PAY",
                "CUMULATIVE_EMPLOYEES_MAIN_SECTION_SCHEME_CONTS",
                "CUMULATIVE_EMPLOYERS_SCHEME_CONTS",
            ):
                val = _to_float(_v(row, fld))
                if val is not None and val < 0:
                    flag(rn, ni, fld, "F031", f"Value: {val:.2f}")

    # Cross-row: PAYROLL_PERIOD_END_DATE must be the same on every row
    if len(period_dates) > 1:
        canonical = next(iter(period_dates))   # first date seen
        for date_val, row_nums in period_dates.items():
            if date_val != canonical:
                for rn in row_nums:
                    row_ni = rows[rn - 2].get("NI_NUMBER", "?") if rn - 2 < len(rows) else "?"
                    flag(rn, row_ni, "PAYROLL_PERIOD_END_DATE", "F008",
                         f"Value: {date_val} — file also contains {canonical}")

    return file_flags, row_flags, rows, headers


# ── Cross-period validation engine ──────────────────────────────────────────────────

def cross_period_checks(period_results):
    """
    period_results: list of dicts, already sorted chronologically by period_date:
        {file, period_date (datetime|None), period_date_str, rows, headers}

    Returns (cross_flags, coverage_rows):
        cross_flags    — list of dicts: file, period, row, ni, field, code, detail
        coverage_rows  — one dict per file: file, period, rows, members_added, members_dropped
    """
    cross_flags   = []
    coverage_rows = []

    def cflag(file, period_str, rnum, ni, field, code, detail=""):
        cross_flags.append(dict(file=file, period=period_str, row=rnum, ni=ni,
                                 field=field, code=code, detail=detail))

    # F035 — gaps between consecutive dated periods
    dated = [p for p in period_results if p["period_date"] is not None]
    for prev, cur in zip(dated, dated[1:]):
        gap = (cur["period_date"] - prev["period_date"]).days
        if gap > MAX_GAP_DAYS:
            cflag(cur["file"], cur["period_date_str"], "FILE", "—", "PAYROLL_PERIOD_END_DATE",
                  "F035",
                  f"{gap} days since the previous period ({prev['period_date_str']} → "
                  f"{cur['period_date_str']}); a submission may be missing")

    member_state = {}   # member key -> tracked state dict
    prev_keys    = None

    for pidx, period in enumerate(period_results):
        file  = period["file"]
        pstr  = period["period_date_str"]
        rows  = period["rows"]

        cur_keys       = set()
        members_added  = 0
        members_dropped = 0

        for ridx, row in enumerate(rows, start=2):
            key = _member_key(row)
            if not key[0]:
                continue   # blank NI — already flagged by the single-file pass
            cur_keys.add(key)
            ni = _v(row, "NI_NUMBER")

            dol    = _v(row, "DATE_OF_LEAVING")
            optout = _v(row, "OPT_OUT_DATE")
            optin  = _v(row, "OPT_IN_DATE")
            djp    = _v(row, "DATE_JOINED_PENSION_SCHEME")
            dob    = str(_v(row, "DOB")).strip()
            gender = str(_v(row, "GENDER")).strip().upper()
            is_50  = _is_5050(row)
            pen_pay     = _to_float(_v(row, "PENSIONABLE_PAY")) or 0
            scheme_rate = _to_float(_v(row, "SCHEME_CONT_RATE"))
            emp_main    = _to_float(_v(row, "EMPLOYEES_MAIN_SECTION_CONTS")) or 0
            emp_50      = _to_float(_v(row, "EMPLOYEES_5050_CONTS")) or 0
            is_refund_row = (not _blank(optout)) and pen_pay == 0

            state = member_state.get(key)

            if state is None:
                members_added += 1
                if pidx > 0 and _blank(djp) and _blank(optin):
                    cflag(file, pstr, ridx, ni, "DATE_JOINED_PENSION_SCHEME", "F037",
                          "Member appears for the first time mid-year with no "
                          "DATE_JOINED_PENSION_SCHEME or OPT_IN_DATE")
            else:
                # Reappeared after previously being marked as a leaver
                if state.get("left") and _blank(djp) and _blank(optin):
                    cflag(file, pstr, ridx, ni, "DATE_JOINED_PENSION_SCHEME", "F043",
                          f"Member had DATE_OF_LEAVING reported in {state.get('left_period')} "
                          f"and reappears with no new join/opt-in date")

                # Static data changes
                if state.get("dob") and dob and state["dob"] != dob:
                    cflag(file, pstr, ridx, ni, "DOB", "F044",
                          f"DOB changed from {state['dob']} to {dob}")
                if state.get("gender") and gender and state["gender"] != gender:
                    cflag(file, pstr, ridx, ni, "GENDER", "F044",
                          f"GENDER changed from {state['gender']} to {gender}")

                prev_cum = state.get("cumulatives", {})

                # Cumulative consistency checks
                if not is_refund_row:
                    for cum_field, period_field, section in CUMULATIVE_PAIRS:
                        if section == "main" and is_50:
                            continue
                        if section == "5050" and not is_50:
                            continue
                        cur_cum = _to_float(_v(row, cum_field))
                        if cur_cum is None:
                            continue
                        prev_val = prev_cum.get(cum_field)
                        if prev_val is None:
                            continue   # no baseline yet for this field
                        per_val = _to_float(_v(row, period_field))
                        if cur_cum < prev_val - 0.02:
                            cflag(file, pstr, ridx, ni, cum_field, "F039",
                                  f"Cumulative fell from {prev_val:.2f} to {cur_cum:.2f} between periods")
                        elif per_val is not None:
                            expected = prev_val + per_val
                            if abs(cur_cum - expected) > 0.02:
                                cflag(file, pstr, ridx, ni, cum_field, "F038",
                                      f"Expected {expected:.2f} ({prev_val:.2f} + {per_val:.2f}), got {cur_cum:.2f}")

                # Post-opt-out activity: opted out previously, not opting in/out again this row
                if state.get("opted_out") and _blank(optout) and _blank(optin):
                    if pen_pay != 0 or emp_main != 0 or emp_50 != 0:
                        cflag(file, pstr, ridx, ni, "PENSIONABLE_PAY", "F040",
                              "New pensionable pay or contributions reported after a prior opt-out")
                    else:
                        for cum_field, _period_field, section in CUMULATIVE_PAIRS:
                            if section == "main" and is_50:
                                continue
                            if section == "5050" and not is_50:
                                continue
                            cur_cum  = _to_float(_v(row, cum_field))
                            prev_val = prev_cum.get(cum_field)
                            if cur_cum is not None and prev_val is not None and abs(cur_cum - prev_val) > 0.02:
                                cflag(file, pstr, ridx, ni, cum_field, "F040",
                                      f"Cumulative changed from {prev_val:.2f} to {cur_cum:.2f} after "
                                      f"opt-out; it should carry forward unchanged")
                    if scheme_rate is not None and scheme_rate != 0:
                        cflag(file, pstr, ridx, ni, "SCHEME_CONT_RATE", "F041",
                              f"SCHEME_CONT_RATE is {scheme_rate}, expected 0 in a period after opt-out")

                # OPT_IN_DATE carried forward beyond the period it was first reported
                if state.get("opt_in_reported") and not _blank(optin):
                    cflag(file, pstr, ridx, ni, "OPT_IN_DATE", "F042",
                          "OPT_IN_DATE still populated in a later period; clear it once the "
                          "opt-in has been processed")

            # ── Roll state forward for the next period ──
            new_state = dict(state) if state else {}
            new_state["dob"]    = dob or new_state.get("dob")
            new_state["gender"] = gender or new_state.get("gender")
            new_state["cumulatives"] = {
                cf: _to_float(_v(row, cf)) for cf, _, _ in CUMULATIVE_PAIRS
                if _to_float(_v(row, cf)) is not None
            }
            if not _blank(dol):
                new_state["left"] = True
                new_state["left_period"] = pstr
            if not _blank(optout) and _blank(optin):
                new_state["opted_out"] = True
            if not _blank(optin):
                new_state["opted_out"]      = False
                new_state["opt_in_reported"] = True
            elif state and state.get("opt_in_reported"):
                new_state["opt_in_reported"] = False   # cleared correctly this period
            if not _blank(djp) or not _blank(optin):
                new_state["left"] = False              # treat as (re)joined
            member_state[key] = new_state

        # Members present last period but missing this period
        if prev_keys is not None:
            for key in (prev_keys - cur_keys):
                state = member_state.get(key, {})
                if not state.get("left"):
                    cflag(file, pstr, "—", key[0], "—", "F036",
                          "Present in the previous period but missing here with no "
                          "DATE_OF_LEAVING recorded")
                    members_dropped += 1

        coverage_rows.append(dict(
            file=file, period=pstr, rows=len(rows),
            members_added=members_added, members_dropped=members_dropped,
        ))
        prev_keys = cur_keys

    return cross_flags, coverage_rows


# ── Excel builder ───────────────────────────────────────────────────────────────────────

def _thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def _style_header(ws, row_idx, fill_hex):
    fill   = PatternFill("solid", fgColor=fill_hex)
    font   = Font(bold=True, color="FFFFFF")
    border = _thin_border()
    align  = Alignment(wrap_text=True, vertical="center", horizontal="center")
    for cell in ws[row_idx]:
        cell.fill      = fill
        cell.font      = font
        cell.border    = border
        cell.alignment = align


def _auto_width(ws, min_w=10, max_w=55):
    for col in ws.columns:
        length = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = \
            min(max(length + 2, min_w), max_w)


# Severity → fill colour
SEV_FILL = {
    "MANDATORY_BLANK":        "FF4C4C",   # red
    "NI_FORMAT":              "FF4C4C",
    "DATE_FORMAT":            "FF4C4C",
    "DATE_INVALID":           "FF9999",   # light-red
    "POSTCODE_FORMAT":        "FFA500",   # amber
    "GENDER_INVALID":         "FF4C4C",
    "NUMERIC_FORMAT":         "FF4C4C",
    "PERIOD_DATE_MISMATCH":   "FF4C4C",
    "FILLER_POPULATED":       "FFA500",
    "COLUMN_COUNT":           "FF4C4C",
    "COLUMN_ORDER":           "FF4C4C",
    "MISSING_CONDITIONAL":    "FFA500",
    "LEAVER_NO_REASON":       "FFA500",
    "BREAK_NO_REASON":        "FFA500",
    "BREAK_REASON_INVALID":   "FFA500",
    "MARITAL_STATUS_INVALID": "FFF2CC",   # yellow
    "PT_INDICATOR_INVALID":   "FFA500",
    "PT_FIELDS_INCONSISTENT": "FFA500",
    "OPT_OUT_AND_IN_SAME_ROW":"FF4C4C",
    "OPT_OUT_RATE_NOT_ZERO":  "FFA500",
    "TAXABLE_EARNINGS_CHECK": "FFA500",
    "CONT_RATE_BAND_CHECK":   "FFA500",
    "5050_MAIN_OVERLAP":      "FFA500",
    "5050_PAY_WRONG_COL":     "FFA500",
    "CUMULATIVE_BELOW_PERIOD":"FFA500",
    "DOL_USED_FOR_OPTOUT":    "FFA500",
    "SURNAME_NOT_CAPS":       "FFF2CC",
    "FORENAMES_NOT_CAPS":     "FFF2CC",
    "TITLE_NOT_CAPS":         "FFF2CC",
    "JOB_TITLE_NOT_CAPS":     "FFF2CC",
    "UNEXPECTED_NEGATIVE":    "FFA500",
    "PAY_REF_TOO_LONG":       "FFA500",
    "MISSING_COLUMN":         "FF4C4C",
    "DOL_AND_OPTOUT_SAME_ROW":"FF4C4C",
    # Cross-period
    "MISSING_PERIOD_GAP":        "FFA500",
    "MEMBER_VANISHED":           "FF4C4C",
    "NEW_MEMBER_NO_JOIN_DATE":   "FFA500",
    "CUMULATIVE_MISMATCH":       "FFA500",
    "CUMULATIVE_DECREASED":      "FF4C4C",
    "POST_OPTOUT_ACTIVITY":      "FF4C4C",
    "POST_OPTOUT_RATE_NOT_ZERO": "FFA500",
    "OPT_IN_CARRIED_FORWARD":    "FFA500",
    "LEAVER_REAPPEARED":         "FFA500",
    "STATIC_DATA_CHANGED":       "FFF2CC",
}


def build_excel(all_flags, cross_flags, coverage_rows, output_path):
    """
    all_flags:   combined single-file flags across every input file, each dict:
                 file, period, row, ni, field, code, detail
    cross_flags: cross-period flags, same shape
    coverage_rows: one dict per file: file, period, rows, members_added, members_dropped
    """
    wb = openpyxl.Workbook()
    border = _thin_border()

    def enrich(f):
        sev, desc = FLAGS[f["code"]]
        return dict(f, sev=sev, desc=desc)

    enriched_flags = [enrich(f) for f in all_flags]
    enriched_cross = [enrich(f) for f in cross_flags]

    # ── Sheet 1: Validation Results ──────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Validation Results"
    ws1.append(["File", "Period", "Row", "NI Number", "Field", "Flag Code",
                "Severity Category", "Description", "Detail"])
    _style_header(ws1, 1, "2E4057")
    ws1.freeze_panes = "A2"

    for f in enriched_flags:
        ws1.append([f["file"], f["period"], f["row"], f["ni"], f["field"], f["code"],
                    f["sev"], f["desc"], f["detail"]])
        r = ws1.max_row
        hex_col = SEV_FILL.get(f["sev"], "FFFFFF")
        fill = PatternFill("solid", fgColor=hex_col)
        for c in range(1, 10):
            ws1.cell(r, c).border = border
        ws1.cell(r, 6).fill = fill   # colour the Flag Code cell

    if not enriched_flags:
        ws1.append(["—", "—", "—", "—", "—", "PASS", "No issues found",
                    "All checks passed for the file(s) supplied.", ""])

    ws1.auto_filter.ref = ws1.dimensions
    _auto_width(ws1)

    # ── Sheet 2: Summary by Flag ─────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary by Flag")
    ws2.append(["Flag Code", "Severity Category", "Description", "Count",
                "Rows Affected (first 20)"])
    _style_header(ws2, 1, "2E4057")
    ws2.freeze_panes = "A2"

    counter = defaultdict(list)
    for f in enriched_flags + enriched_cross:
        counter[f["code"]].append(f"{f['file']}:{f['row']}")

    for code in sorted(counter):
        rows_aff = counter[code]
        sev, desc = FLAGS.get(code, ("?", "?"))
        sample = ", ".join(rows_aff[:20]) + ("…" if len(rows_aff) > 20 else "")
        ws2.append([code, sev, desc, len(rows_aff), sample])
        r = ws2.max_row
        for c in range(1, 6):
            ws2.cell(r, c).border = border

    if not counter:
        ws2.append(["—", "—", "No issues found", 0, "—"])

    _auto_width(ws2)

    # ── Sheet 3: Cross-Period Checks ─────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Cross-Period Checks")
    ws3.append(["File", "Period", "Row", "NI Number", "Field", "Flag Code",
                "Severity Category", "Description", "Detail"])
    _style_header(ws3, 1, "2E4057")
    ws3.freeze_panes = "A2"

    for f in enriched_cross:
        ws3.append([f["file"], f["period"], f["row"], f["ni"], f["field"], f["code"],
                    f["sev"], f["desc"], f["detail"]])
        r = ws3.max_row
        hex_col = SEV_FILL.get(f["sev"], "FFFFFF")
        fill = PatternFill("solid", fgColor=hex_col)
        for c in range(1, 10):
            ws3.cell(r, c).border = border
        ws3.cell(r, 6).fill = fill

    if not enriched_cross:
        note = ("No cross-period issues found." if len(coverage_rows) > 1
                else "Only one period was supplied — cross-period checks need 2+ files.")
        ws3.append(["—", "—", "—", "—", "—", "—", "—", note, ""])

    ws3.auto_filter.ref = ws3.dimensions
    _auto_width(ws3)

    # ── Sheet 4: Period Coverage ─────────────────────────────────────────────────────────
    ws4 = wb.create_sheet("Period Coverage")
    ws4.append(["File", "Period End Date", "Rows", "Members Added", "Members Dropped"])
    _style_header(ws4, 1, "375623")
    ws4.freeze_panes = "A2"

    for c in coverage_rows:
        ws4.append([c["file"], c["period"], c["rows"], c["members_added"], c["members_dropped"]])
        r = ws4.max_row
        for col in range(1, 6):
            ws4.cell(r, col).border = border

    if not coverage_rows:
        ws4.append(["—", "—", 0, 0, 0])

    _auto_width(ws4)

    # ── Sheet 5: Flag Reference ────────────────────────────────────────────────────────────
    ws5 = wb.create_sheet("Flag Reference")
    ws5.append(["Flag Code", "Severity Category", "Short Description",
                "How to Fix / Guide Reference"])
    _style_header(ws5, 1, "375623")
    ws5.freeze_panes = "A2"

    for code in sorted(FLAGS):
        sev, desc = FLAGS[code]
        fix = FIX_GUIDANCE.get(code, "")
        ws5.append([code, sev, desc, fix])
        r = ws5.max_row
        hex_col = SEV_FILL.get(sev, "FFFFFF")
        fill = PatternFill("solid", fgColor=hex_col)
        ws5.cell(r, 1).fill = fill
        for c in range(1, 5):
            ws5.cell(r, c).border = border
            ws5.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")

    _auto_width(ws5)
    ws5.column_dimensions["D"].width = 75

    # ── Sheet 6: Colour Key ──────────────────────────────────────────────────────────────
    ws6 = wb.create_sheet("Colour Key")
    ws6.append(["Colour", "Meaning", "Example Flag Codes"])
    _style_header(ws6, 1, "2E4057")
    key_rows = [
        ("FF4C4C", "HIGH — file will likely be rejected / data or compliance risk",
         "F001, F002, F003, F007, F008, F010, F011, F033, F034, F036, F039, F040"),
        ("FFA500", "MEDIUM — incorrect data that affects member records",
         "F005, F009, F012, F013, F021, F022, F023, F025, F031, F035, F037, F038, F041, F042, F043"),
        ("FFF2CC", "LOW — formatting issue that should be corrected",
         "F016, F027, F028, F029, F030, F044"),
    ]
    for hex_col, meaning, examples in key_rows:
        ws6.append(["", meaning, examples])
        r = ws6.max_row
        ws6.cell(r, 1).fill = PatternFill("solid", fgColor=hex_col)
        for c in range(1, 4):
            ws6.cell(r, c).border = border
    _auto_width(ws6)

    # ── Sheet 7: Contribution Rates 2026/27 ─────────────────────────────────────────────
    ws7 = wb.create_sheet("Contribution Rates 2026-27")
    ws7.append(["Band", "Actual Pensionable Pay — From (£)", "To (£)",
                "Main Section Rate", "50/50 Section Rate"])
    _style_header(ws7, 1, "375623")

    band_data = [
        (1, "0.00",          "18,400.00",   "5.50%", "2.75%"),
        (2, "18,400.01",     "29,000.00",   "5.80%", "2.90%"),
        (3, "29,000.01",     "47,300.00",   "6.50%", "3.25%"),
        (4, "47,300.01",     "59,800.00",   "6.80%", "3.40%"),
        (5, "59,800.01",     "84,000.00",   "8.50%", "4.25%"),
        (6, "84,000.01",     "119,100.00",  "9.90%", "4.95%"),
        (7, "119,100.01",    "140,400.00",  "10.50%","5.25%"),
        (8, "140,400.01",    "210,700.00",  "11.40%","5.70%"),
        (9, "210,700.01",    "No upper limit","12.50%","6.25%"),
    ]
    for bd in band_data:
        ws7.append(list(bd))
        for c in range(1, 6):
            ws7.cell(ws7.max_row, c).border = border

    ws7.append([])
    note = ("Notes:\n"
            "• Rates apply 1 April 2026 – 31 March 2027.\n"
            "• Use the member's ACTUAL pensionable pay (not FTE) to determine their band.\n"
            "• Part-time members: compare their actual contracted pay to the bands, then record the matching rate in SCHEME_CONT_RATE.\n"
            "• The 50/50 rate is exactly half the main-section rate.\n"
            "• Thresholds reflect a 3.8% CPI uplift from 2025/26; percentage rates are unchanged.")
    ws7.append([note])
    ws7.cell(ws7.max_row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws7.row_dimensions[ws7.max_row].height = 90
    ws7.merge_cells(f"A{ws7.max_row}:E{ws7.max_row}")
    _auto_width(ws7)

    wb.save(output_path)


# ── Main ─────────────────────────────────────────────────────────────────────────────────

def _parse_args():
    parser = argparse.ArgumentParser(
        description="Validate i-Connect payroll extract CSV(s) against the Enfield "
                    "Pension Fund 2026/27 File Completion Guide.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python iconnect_validator.py april.csv\n"
            "  python iconnect_validator.py april.csv may.csv june.csv -o report.xlsx\n"
            "  python iconnect_validator.py --dir ./extracts/ -o report.xlsx\n"
        ),
    )
    parser.add_argument("files", nargs="*", help="One or more CSV extract files")
    parser.add_argument("--dir", dest="directory",
                         help="Directory containing CSV extracts (all *.csv files are used)")
    parser.add_argument("-o", "--output", dest="output",
                         help="Output .xlsx path (default: derived from the first input file)")
    return parser.parse_args()


def main():
    args = _parse_args()

    csv_paths = list(args.files)
    if args.directory:
        found = sorted(glob.glob(os.path.join(args.directory, "*.csv")))
        if not found:
            sys.exit(f"ERROR: No .csv files found in {args.directory}")
        csv_paths.extend(found)

    if not csv_paths:
        print(__doc__)
        sys.exit(0)

    for p in csv_paths:
        if not os.path.isfile(p):
            sys.exit(f"ERROR: File not found: {p}")

    output_path = args.output or (
        os.path.splitext(csv_paths[0])[0] + "_validation_flags.xlsx"
        if len(csv_paths) == 1 else "iconnect_validation_flags.xlsx"
    )

    # ── Validate each file independently ──
    period_results = []
    all_flags = []

    for path in csv_paths:
        print(f"Validating: {path}")
        file_flags, row_flags, rows, headers = validate(path)
        fname = os.path.basename(path)
        period_date     = _first_valid_period_date(rows)
        period_date_str = period_date.strftime("%d/%m/%Y") if period_date else "Unknown"

        for f in file_flags + row_flags:
            all_flags.append(dict(f, file=fname, period=period_date_str))

        period_results.append(dict(
            file=fname, period_date=period_date, period_date_str=period_date_str,
            rows=rows, headers=headers,
        ))
        print(f"  Rows: {len(rows)}  |  Period: {period_date_str}  |  "
              f"Flags: {len(file_flags) + len(row_flags)}")

    # ── Order periods chronologically; undated periods sort last with a warning ──
    dated   = [p for p in period_results if p["period_date"] is not None]
    undated = [p for p in period_results if p["period_date"] is None]
    dated.sort(key=lambda p: p["period_date"])
    if undated:
        print(f"\nWARNING: {len(undated)} file(s) have no readable PAYROLL_PERIOD_END_DATE "
              f"and could not be placed in chronological order: "
              f"{', '.join(p['file'] for p in undated)}")
    period_results = dated + undated

    # ── Reject duplicate period end dates across files (ambiguous submissions) ──
    seen_dates = defaultdict(list)
    for p in dated:
        seen_dates[p["period_date_str"]].append(p["file"])
    dupes = {d: files for d, files in seen_dates.items() if len(files) > 1}
    if dupes:
        print("\nERROR: More than one file claims the same PAYROLL_PERIOD_END_DATE:")
        for d, files in dupes.items():
            print(f"  {d}: {', '.join(files)}")
        sys.exit("Resolve the duplicate submission(s) before validating.")

    # ── Cross-period pass (only meaningful with 2+ files, but safe to always run) ──
    cross_flags, coverage_rows = cross_period_checks(period_results)

    total_single = len(all_flags)
    total_cross  = len(cross_flags)

    build_excel(all_flags, cross_flags, coverage_rows, output_path)

    print(f"\nFiles validated: {len(csv_paths)}")
    print(f"Single-file flags: {total_single}")
    if len(csv_paths) > 1:
        print(f"Cross-period flags: {total_cross}")
    print(f"\nOutput: {output_path}")
    print("Sheets:")
    print("  1. Validation Results   — every within-file flag: file, period, row, field, code")
    print("  2. Summary by Flag      — count and rows affected per flag code (all flags)")
    print("  3. Cross-Period Checks  — flags that only exist by comparing periods")
    print("  4. Period Coverage      — rows/joiners/leavers per file, in chronological order")
    print("  5. Flag Reference       — all 44 flag codes with fix guidance")
    print("  6. Colour Key           — what the highlight colours mean")
    print("  7. Contribution Rates   — 2026/27 band table for SCHEME_CONT_RATE")

    grand_total = total_single + total_cross
    if grand_total == 0:
        print("\nNo issues found — file(s) appear to conform to the guide.")
    else:
        high = sum(1 for f in (all_flags + cross_flags)
                   if SEV_FILL.get(FLAGS[f["code"]][0], "") == "FF4C4C")
        med  = grand_total - high
        print(f"\n  HIGH (red)   : {high}")
        print(f"  MEDIUM/LOW   : {med}")


if __name__ == "__main__":
    main()
