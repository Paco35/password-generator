#!/usr/bin/env python3
"""
iConnect Payroll Extract Validator — Enfield Pension Fund 2026/27
Validates a CSV payroll extract against the File Completion Guide and
writes all flags to an Excel workbook.

Usage:
    python iconnect_validator.py <extract.csv> [output.xlsx]
"""

import csv
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


# ── Validation engine ───────────────────────────────────────────────────────────────

def validate(csv_path):
    """Return (file_flags, row_flags, rows, actual_headers)."""

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
    period_dates: dict[str, list[int]] = defaultdict(list)

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
}


def build_excel(file_flags, row_flags, rows, output_path, source_filename):
    wb = openpyxl.Workbook()
    border = _thin_border()

    all_flags = [
        dict(row="FILE", ni="—", field=f["field"],
             code=f["code"],
             sev=FLAGS[f["code"]][0],
             desc=FLAGS[f["code"]][1],
             detail=f["detail"])
        for f in file_flags
    ] + [
        dict(row=f["row"], ni=f["ni"], field=f["field"],
             code=f["code"],
             sev=FLAGS[f["code"]][0],
             desc=FLAGS[f["code"]][1],
             detail=f["detail"])
        for f in row_flags
    ]

    # ── Sheet 1: Validation Results ──────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Validation Results"
    ws1.append(["Row", "NI Number", "Field", "Flag Code",
                "Severity Category", "Description", "Detail"])
    _style_header(ws1, 1, "2E4057")
    ws1.freeze_panes = "A2"

    for f in all_flags:
        ws1.append([f["row"], f["ni"], f["field"], f["code"],
                    f["sev"], f["desc"], f["detail"]])
        r = ws1.max_row
        hex_col = SEV_FILL.get(f["sev"], "FFFFFF")
        fill = PatternFill("solid", fgColor=hex_col)
        for c in range(1, 8):
            ws1.cell(r, c).border = border
        ws1.cell(r, 4).fill = fill   # colour the Flag Code cell

    if not all_flags:
        ws1.append(["—", "—", "—", "PASS", "No issues found",
                    "All checks passed for this file.", ""])

    ws1.auto_filter.ref = ws1.dimensions
    _auto_width(ws1)

    # ── Sheet 2: Summary by Flag ─────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary by Flag")
    ws2.append(["Flag Code", "Severity Category", "Description", "Count",
                "Rows Affected (first 20)"])
    _style_header(ws2, 1, "2E4057")
    ws2.freeze_panes = "A2"

    counter = defaultdict(list)
    for f in all_flags:
        counter[f["code"]].append(str(f["row"]))

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

    # ── Sheet 3: Flag Reference ────────────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Flag Reference")
    ws3.append(["Flag Code", "Severity Category", "Short Description",
                "How to Fix / Guide Reference"])
    _style_header(ws3, 1, "375623")
    ws3.freeze_panes = "A2"

    for code in sorted(FLAGS):
        sev, desc = FLAGS[code]
        fix = FIX_GUIDANCE.get(code, "")
        ws3.append([code, sev, desc, fix])
        r = ws3.max_row
        hex_col = SEV_FILL.get(sev, "FFFFFF")
        fill = PatternFill("solid", fgColor=hex_col)
        ws3.cell(r, 1).fill = fill
        for c in range(1, 5):
            ws3.cell(r, c).border = border
            ws3.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")

    _auto_width(ws3)
    ws3.column_dimensions["D"].width = 75

    # ── Sheet 4: Colour Key ──────────────────────────────────────────────────────────────
    ws4 = wb.create_sheet("Colour Key")
    ws4.append(["Colour", "Meaning", "Example Flag Codes"])
    _style_header(ws4, 1, "2E4057")
    key_rows = [
        ("FF4C4C", "HIGH — file will likely be rejected / data loss risk",
         "F001, F002, F003, F007, F008, F010, F011, F033, F034"),
        ("FFA500", "MEDIUM — incorrect data that affects member records",
         "F005, F009, F012, F013, F021, F022, F023, F025, F031"),
        ("FFF2CC", "LOW — formatting issue that should be corrected",
         "F016, F027, F028, F029, F030"),
    ]
    for hex_col, meaning, examples in key_rows:
        ws4.append(["", meaning, examples])
        r = ws4.max_row
        ws4.cell(r, 1).fill = PatternFill("solid", fgColor=hex_col)
        for c in range(1, 4):
            ws4.cell(r, c).border = border
    _auto_width(ws4)

    # ── Sheet 5: Contribution Rates 2026/27 ─────────────────────────────────────────────
    ws5 = wb.create_sheet("Contribution Rates 2026-27")
    ws5.append(["Band", "Actual Pensionable Pay — From (£)", "To (£)",
                "Main Section Rate", "50/50 Section Rate"])
    _style_header(ws5, 1, "375623")

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
        ws5.append(list(bd))
        for c in range(1, 6):
            ws5.cell(ws5.max_row, c).border = border

    ws5.append([])
    note = ("Notes:\n"
            "• Rates apply 1 April 2026 – 31 March 2027.\n"
            "• Use the member's ACTUAL pensionable pay (not FTE) to determine their band.\n"
            "• Part-time members: compare their actual contracted pay to the bands, then record the matching rate in SCHEME_CONT_RATE.\n"
            "• The 50/50 rate is exactly half the main-section rate.\n"
            "• Thresholds reflect a 3.8% CPI uplift from 2025/26; percentage rates are unchanged.")
    ws5.append([note])
    ws5.cell(ws5.max_row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws5.row_dimensions[ws5.max_row].height = 90
    ws5.merge_cells(f"A{ws5.max_row}:E{ws5.max_row}")
    _auto_width(ws5)

    wb.save(output_path)


# ── Main ─────────────────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("Example:")
        print("  python iconnect_validator.py payroll_april2026.csv")
        print("  python iconnect_validator.py payroll_april2026.csv my_report.xlsx")
        sys.exit(0)

    csv_path = sys.argv[1]
    if not os.path.isfile(csv_path):
        sys.exit(f"ERROR: File not found: {csv_path}")

    output_path = sys.argv[2] if len(sys.argv) >= 3 else \
        os.path.splitext(csv_path)[0] + "_validation_flags.xlsx"

    print(f"Validating: {csv_path}")
    file_flags, row_flags, rows, headers = validate(csv_path)

    total = len(file_flags) + len(row_flags)
    print(f"Rows read:  {len(rows)}")
    print(f"Flags found: {total}  "
          f"(file-level: {len(file_flags)}, row-level: {len(row_flags)})")

    build_excel(file_flags, row_flags, rows, output_path,
                os.path.basename(csv_path))

    print(f"\nOutput: {output_path}")
    print("Sheets:")
    print("  1. Validation Results   — every flag with row, NI, field, code, detail")
    print("  2. Summary by Flag      — count and rows affected per flag code")
    print("  3. Flag Reference       — all 34 flag codes with fix guidance")
    print("  4. Colour Key           — what the highlight colours mean")
    print("  5. Contribution Rates   — 2026/27 band table for SCHEME_CONT_RATE")

    if total == 0:
        print("\nNo issues found — file appears to conform to the guide.")
    else:
        high = sum(1 for f in (file_flags + row_flags)
                   if SEV_FILL.get(FLAGS[f["code"]][0], "") == "FF4C4C")
        med  = total - high
        print(f"\n  HIGH (red)   : {high}")
        print(f"  MEDIUM/LOW   : {med}")


if __name__ == "__main__":
    main()
