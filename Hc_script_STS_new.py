"""
Hc_script_STS_new.py

Merge PII/PHI person records from an Excel export into one row per
confirmed person, for the notification report.

This is a restructured version of the prior merge logic (see
hc_script_old.py / hc_script_unknowns.py for the earlier design). The
pipeline is now built as three sequential phases instead of one flat set of
8 priority levels:

  PHASE 1 - Unknown Name Split (Steps 1-2 of the new spec)
    Any row missing a usable First Name and/or Last Name is pulled OUT of
    the merge pool entirely, before any matching happens, and written to a
    separate "Unknown_Entries" sheet - it is NEVER compared against other
    rows or merged with anything (a name-less/partial-name row is not
    trustworthy enough to identity-match on ID alone in this version). In
    that sheet, whichever of First/Last Name was blank is displayed as
    "[Unknown]" (the other side keeps its real value if it had one); an
    "Unknown Field" column says which side(s) were blank. This is a raw,
    UNMERGED list - rows here are not deduped against each other, so a
    manual reviewer sees every original row untouched.

  PHASE 2 - Name+Suffix ID Cascade (Step 3 of the new spec)
    Every remaining row (real First AND Last Name) is first partitioned by
    an EXACT normalized (First Name, Last Name, Suffix) key - two rows can
    ONLY ever end up in the same merged person if they share this exact
    key. (This is a deliberate change from the old script's prefix/initial-
    tolerant name matching - see the "Name grouping" note below.) WITHIN
    each such name partition, rows are further clustered by 5 strict
    priority levels, applied in order and LOCKED same as the old script's
    LEVEL_ORDER mechanism (a level's groups, once formed, can gain new rows
    from a later/weaker level but can never be bridged into another already-
    locked group by a later level):
        Level 1: SSN                    (Step 3.1)
        Level 2: SSN + DOB              (Step 3.2)
        Level 3: Driver's License       (Step 3.3)
        Level 4: Passport               (Step 3.4)
        Level 5: Tax Identification     (Step 3.5)
    Since every pair compared already shares the exact same Name+Suffix,
    Name itself is no longer part of any level's match test - each level
    only checks the ID evidence itself, plus the same conflict guards as
    before (a genuinely differing DOB blocks a same-SSN match; a genuinely
    differing SSN blocks a same-DOB match; a genuinely differing Middle
    Name blocks any level - blank/prefix-compatible Middle Name is fine).
    Employee ID, Phone Number, Email Address, and every address field are
    NO LONGER matching criteria in this version (the old script's Rules 4,
    7, 8, 9 are gone) - they are purely APPENDED attributes now (Step 4):
    every value seen across a merged group's rows is kept, semicolon-
    joined, regardless of which level actually matched. Address specifically
    keeps the OLD majority-address / "Other Address" logic (Step 4.1) - see
    split_addresses() - since a blank address should never spin off a
    spurious extra address, and a group can still have have gathered rows
    with 2+ genuinely different real addresses.

    "Tax Identification" (Step 3.5) matches on COL_TAXID ("Tax
    Identification Number"), a real, distinct column in the source data -
    Government-Issued ID Number (COL_GOVID) remains a separate, purely
    APPEND-ONLY field (see OTHER_MERGE_COLS), same as it always was.

    Two source columns are handled specially, outside the 5-level cascade:
      - COL_DOB_ALT ("DOBs", plural) - a secondary/legacy DOB source,
        confirmed to feed the SAME date evidence as COL_DOB itself, not an
        independent field. Its value is used for matching (Steps 3.1/3.2)
        whenever COL_DOB itself is blank, and its raw value is folded into
        the same dob_merge() call that builds the output DOB display - it
        never appears as its own output column.
      - COL_UNIQUE_ID ("Unique_ID") - the source database's own primary
        key for a row, confirmed NOT to be a person-identity signal (never
        used for matching). The output simply keeps whichever value was on
        the FIRST (topmost) original row in a merged group and drops every
        other row's value - unlike every OTHER_MERGE_COLS field, it is
        never semicolon-merged (see main()'s Step 4 build loop and
        _fold_row_into(), which both deliberately exclude it).

    ASSUMPTION - Name grouping is an EXACT normalized match (no prefix/
    initial tolerance), taken from the literal "First + Last Name + Suffix
    combination" wording in the spec. If two rows for the same person use
    an initial vs. a full name ("J" vs "Jeffrey"), they will NOT be grouped
    together in this version, unlike the old script's name_prefix_compat().

  PHASE 3 - Blank-ID Fold (Step 5 of the new spec)
    Once Phase 2 is done, each Name+Suffix partition may still hold 2+
    separate output rows (only when nothing bridged them at any of the 5
    levels). These are split into:
      - FILLED rows: at least one of SSN/DOB/Driver's License/Passport/Tax
        ID is populated.
      - BLANK-ID rows: none of those five fields are populated (only a
        name, and maybe Address/Employee ID/Phone/Email/DOCID).
    A BLANK-ID row is folded (its Address/Employee ID/Phone/Email/DOCID
    appended in) into:
      - the partition's one FILLED row, if there is exactly one;
      - the richest of 2+ FILLED rows (most of the 5 ID fields populated),
        if they're unambiguous - flagged True in the "Blank-ID Row Merged
        (Tie-Break)" output column - or left standing alone and reported on
        the "Ambiguous Name-Group Review" sheet if even that's tied;
      - each other (all folded into one row), if the partition has NO
        filled row at all - Name+Suffix alone is the only evidence there
        is, so nothing is left unmerged in that case.
    ASSUMPTION: the tie-break logic and its review sheet reuse the same
    "richest profile wins, true ties get reported" approach the prior
    Unknown Name Bridge pass used (see hc_script_unknowns.py) - Address/
    Employee ID/Phone/Email are NEVER used to decide a fold, only appended
    after the decision is made, consistent with Step 4's append-only role
    for those fields.

INPUT  : a CSV file (INPUT_CSV below) with the columns listed in
         EXPECTED_COLS below - every column read and written as plain text
         (dtype=str), so a leading zero in a ZIP/ID or a masked SSN like
         '123-45-XXXX' is never silently reinterpreted as a number.
OUTPUT : SIX separate CSV files (CSV has no concept of multiple sheets in
         one file, so each of the old workbook's sheets is now its own
         file - see _sheet_csv_path()): OUTPUT_CSV itself is "Merged
         Notification Data"; the other five are written alongside it as
         "<OUTPUT_CSV base name> - <sheet name>.csv" -
         "STS notification merge output - Unknown_Entries.csv", etc.
         - "Merged Notification Data" (OUTPUT_CSV): ONE ROW PER CONFIRMED
           PERSON (known-name rows only - see Unknown_Entries below).
           - First Name, Middle Name, Last Name, Suffix, and SSN: the
             single fullest/most complete value among the merged rows.
           - DOB: every distinct real date seen, "; "-joined, formatted
             MM/DD/YYYY (see dob_merge()).
           - Driver's License, Passport Number, Government-Issued ID
             Number (Tax ID), Employee ID: every distinct ID token seen,
             deduplicated and "; "-joined (cells may already contain
             multiple semicolon-joined tokens - see parse_id_tokens()).
           - Every OTHER column (DOCIDs, Phone, Email, etc.): every
             distinct value seen, "; "-joined. DOCIDs still spills into
             "DOCIDs 2", "DOCIDs 3", ... columns past DOCID_CHUNK_SIZE
             characters (see split_docid_chunks()) - CSV itself has no
             per-cell limit, but this output is routinely opened in Excel,
             which does (32,767 chars/cell), so the safeguard stays.
           - Address fields: the majority address stays in the normal
             columns (gaps filled from another row's fuller copy of the
             SAME address); every other distinct address goes into "Other
             Address", semicolon-joined (see split_addresses()).
           - "Rows Merged": how many original input rows this became.
           - "Names Differ": True if 2+ raw (pre-normalization) First or
             Last Name spellings were seen in the group.
           - "Blank-ID Row Merged (Tie-Break)": True only for a row that
             absorbed a blank-ID row via an evidence tie-break among 2+
             equally-matched filled candidates (see Phase 3) - worth a
             quick double-check.
         - "Unknown_Entries": every raw row pulled out in Phase 1 (First
           and/or Last Name blank/placeholder) - unmerged, one row each,
           with the blank name side(s) displayed as "[Unknown]".
         - "Junk SSN Review": every row where a non-blank SSN value was
           ignored as unusable (a masked SSN with too few known digits).
         - "Junk DOB Review": every row where a non-blank DOB value failed
           to parse (so it was treated as blank instead of silently
           dropped).
         - "Large Group Review": every merged group with more than 50 rows.
         - "Ambiguous Name-Group Review": every blank-ID row that had 2+
           equally-matched filled-row candidates in Phase 3, still tied on
           evidence - left unmerged (standing alone in Merged Notification
           Data) rather than guessed at.

This script does not touch the input file. Save the output only to the
secured/authorized folder for this data (never a desktop) - it contains
SSN, DOB, and other PII/PHI.

Designed for large row counts (uses "blocking" - only compares rows that
already share an exact Name+Suffix key AND an exact SSN, DOB, Driver's
License, Passport, or Tax ID - instead of comparing every row to every
other row).

Install once:
    pip install pandas numpy

Run:
    python Hc_script_STS_new.py
"""

import sys
import os
import re
import time
import itertools
import unicodedata
from collections import defaultdict
from multiprocessing import Pool

import numpy as np
import pandas as pd

# Worker processes for the pairwise clustering step - see hc_script_old.py's
# equivalent comment; plain `threading` would not help here (CPU-bound, GIL-
# serialized), separate processes actually parallelize it.
PARALLEL_WORKERS = max(1, (os.cpu_count() or 4) - 1)
PARALLEL_THRESHOLD = 20_000

# ------------------------------------------------------------
# Progress bar - one line, updated in place, no per-item explanation.
# ------------------------------------------------------------
_last_pct = {}
_label_start = {}


def _format_duration(seconds: float) -> str:
    """'45s', '3m12s', or '1h05m' - whichever is most readable for the size."""
    seconds = max(0, int(seconds))
    if seconds < 60:
        return f"{seconds}s"
    minutes, secs = divmod(seconds, 60)
    if minutes < 60:
        return f"{minutes}m{secs:02d}s"
    hours, minutes = divmod(minutes, 60)
    return f"{hours}h{minutes:02d}m"


def progress(label: str, current: int, total: int, extra: str = "") -> None:
    """Prints '[label]  42% |########------------|  420,000/1,000,000  ETA 3m12s  extra'
    on a single line, overwriting itself."""
    now = time.monotonic()
    start = _label_start.setdefault(label, now)
    pct = 100 if total <= 0 else min(100, current * 100 // total)
    if _last_pct.get(label) == pct and current != total:
        return
    _last_pct[label] = pct
    bar = "#" * (pct // 5) + "-" * (20 - pct // 5)
    elapsed = now - start
    if current >= total:
        timing = f"  (took {_format_duration(elapsed)})"
    elif current > 0:
        remaining = elapsed * (total - current) / current
        timing = f"  ETA {_format_duration(remaining)}"
    else:
        timing = ""
    tail = f"  {extra}" if extra else ""
    end = "\n" if current >= total else ""
    print(f"\r  [{label}] {pct:3d}% |{bar}| {current:,}/{total:,}{timing}{tail}   ",
          end=end, flush=True)


# ------------------------------------------------------------
# 1) CONFIG - edit these to match your workbook
# ------------------------------------------------------------
INPUT_CSV  = "Cng Notification_Final_updated.csv"
OUTPUT_CSV = "STS notification merge output.csv"

COL_DOCID  = "DOCIDs"
COL_FIRST  = "First Name"
COL_LAST   = "Last Name"
COL_MIDDLE = "Middle Name"
COL_SUFFIX = "Suffix"
COL_DOB    = "Full Date of Birth (MM/DD/YYYY)"
COL_SSN    = "Social Security Number"

COL_DL       = "Driver's License Number"
COL_PASSPORT = "Passport Number"
COL_GOVID    = "Government-Issued ID Number"
COL_EMPID    = "Employee Identification Number"
COL_PHONE    = "Phone Number - Personal"
COL_EMAIL    = "Email Address - Personal"
COL_TAXID    = "Tax Identification Number"

# Tax Identification match key (Step 3.5) - a real column in the source
# data (confirmed against the actual export's headers), not the
# Government-Issued ID Number workaround this used before that was known.
TAXID_COL = COL_TAXID

# "DOBs" (plural) - a secondary/legacy DOB source column, confirmed to be
# used the same way as COL_DOB itself, not an independent field: its values
# feed the SAME date evidence used for both matching (Step 3.1/3.2's DOB
# checks) and the output DOB display, and it is NOT a separate output
# column - see build_records()/main()'s dob_merge() calls.
COL_DOB_ALT = "DOBs"

# "Unique_ID" - the source database's own primary key for a row, confirmed
# NOT to be a person-identity signal (rows sharing it aren't necessarily
# the same confirmed person) - never used for matching. Per row, the
# MERGED output simply keeps whichever value was on the first (topmost)
# original row in that merged group and drops every other row's value -
# never semicolon-merged like OTHER_MERGE_COLS - see main()'s Step 4
# build loop and _fold_row_into() (deliberately excluded from both).
COL_UNIQUE_ID = "Unique_ID"

COL_ADDR    = "Residential Address"
COL_CITY    = "City"
COL_STATE   = "State of Residence (if US)"
COL_PROVINCE = "Province of Residence (if Canada)"
COL_ZIP     = "Zip Code"
COL_COUNTRY = "Country of Residence"

ADDRESS_COLS = [COL_ADDR, COL_CITY, COL_STATE, COL_PROVINCE, COL_ZIP, COL_COUNTRY]

# Every other column in the sheet - these get semicolon-merged as-is,
# regardless of which level (if any) matched. Driver's License, Passport,
# Government-Issued ID, Tax ID, Employee ID, Phone, and Email are all
# APPEND-ONLY here now (see Phase 2 in the module docstring) - only SSN,
# DOB, Driver's License, Passport, and Tax ID also double as match keys.
OTHER_MERGE_COLS = [
    "Data Subject Type",
    "Birth Information",
    "Address Comments",
    COL_EMAIL,
    COL_PHONE,
    "Contact Information",
    COL_DL,
    "DL Issuing Country",
    "DL Issuing Province (if Canada)",
    "DL Issuing State (if US)",
    "Passport Country",
    COL_PASSPORT,
    "Government ID Issuing Country",
    "Government- Issued Identification",
    COL_GOVID,
    COL_TAXID,
    "Health Related Information",
    COL_EMPID,
    "Work-Related Information",
    "Family Information",
    "Financial Account Information",
    "Demographic Information",
    "Biometric Data",
    "PI Notes",
    "Access Credentials (Non-Financial Account)",
]

EXPECTED_COLS = (
    [COL_DOCID, COL_FIRST, COL_LAST, COL_MIDDLE, COL_SUFFIX, COL_DOB, COL_DOB_ALT,
     COL_SSN, COL_UNIQUE_ID]
    + ADDRESS_COLS + OTHER_MERGE_COLS
)

MERGE_SEP = "; "

NAME_PLACEHOLDERS = {
    "UNKNOWN", "UNK", "UNKN", "NA", "NONE", "NULL", "NIL",
    "XXX", "XX", "X", "NMN", "NONAME", "NOTGIVEN", "NOTPROVIDED",
}

# Display placeholder for a row pulled into Unknown_Entries (Phase 1) - the
# literal text requested for Step 1/2. "UNKNOWN" is already in
# NAME_PLACEHOLDERS above, so this value is itself always treated as blank/
# non-conflicting by every name-matching function - it never accidentally
# becomes real name evidence if it somehow ended up back in the merge pool.
NAME_UNKNOWN = "[Unknown]"

# ------------------------------------------------------------
# 2) Normalization helpers (unchanged from the prior script - generic text/
#    ID/date/address normalization, none of this needed to change for the
#    new 3-phase structure)
# ------------------------------------------------------------
_INVISIBLE_CODEPOINTS = (0x200B, 0x200C, 0x200D, 0x2060, 0xFEFF, 0x00A0, 0x00AD)
_INVISIBLE_CHARS = re.compile("[" + "".join(chr(c) for c in _INVISIBLE_CODEPOINTS) + "]")


def _numeric_cell_to_str(v) -> str:
    """Coerces a raw cell value to text without letting pandas' float parsing
    corrupt a whole-number ID (SSN, Employee ID, Driver's License, Passport,
    Phone, Tax ID)."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def norm_text(v) -> str:
    """DEDUP KEY everywhere (semicolon_merge, address_key, name matching).
    NFKC-normalizes, strips invisible characters, collapses whitespace,
    upper-cases.

    PERFORMANCE: NFKC-normalize and the invisible-character strip are only
    ever able to change a string that contains a NON-ASCII character - NFKC
    is the identity transform on pure ASCII text, and every codepoint in
    _INVISIBLE_CHARS (zero-width space/joiner, BOM, NBSP, soft hyphen) is
    itself non-ASCII, so neither step can do anything to a pure-ASCII
    string. Real-world name/address data is overwhelmingly ASCII, and
    str.isascii() is a cheap C-level scan, so skipping both (comparatively
    expensive) steps whenever it's True is a safe, byte-identical fast path
    - this function runs on every cell of every row, so it matters."""
    if v is None:
        return ""
    s = str(v)
    if not s.isascii():
        s = unicodedata.normalize("NFKC", s)
        s = _INVISIBLE_CHARS.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip().upper()
    return "" if s.lower() in ("nan", "none", "null") else s


_STREET_TOKEN_MAP = {
    "NORTH": "N", "N": "N", "SOUTH": "S", "S": "S", "EAST": "E", "E": "E",
    "WEST": "W", "W": "W", "NORTHEAST": "NE", "NE": "NE", "NORTHWEST": "NW", "NW": "NW",
    "SOUTHEAST": "SE", "SE": "SE", "SOUTHWEST": "SW", "SW": "SW",
    "STREET": "ST", "ST": "ST", "AVENUE": "AVE", "AVE": "AVE", "AV": "AVE",
    "BOULEVARD": "BLVD", "BLVD": "BLVD", "ROAD": "RD", "RD": "RD",
    "LANE": "LN", "LN": "LN", "DRIVE": "DR", "DR": "DR", "COURT": "CT", "CT": "CT",
    "CIRCLE": "CIR", "CIR": "CIR", "PLACE": "PL", "PL": "PL",
    "TERRACE": "TER", "TERR": "TER", "TER": "TER", "PARKWAY": "PKWY", "PKWY": "PKWY",
    "HIGHWAY": "HWY", "HWY": "HWY", "SQUARE": "SQ", "SQ": "SQ",
    "TRAIL": "TRL", "TRL": "TRL", "WAY": "WAY", "LOOP": "LOOP",
    "COVE": "CV", "CV": "CV", "POINT": "PT", "PT": "PT",
    "CROSSING": "XING", "XING": "XING", "PLAZA": "PLZ", "PLZ": "PLZ",
    "EXPRESSWAY": "EXPY", "EXPY": "EXPY", "FREEWAY": "FWY", "FWY": "FWY",
    "ROUTE": "RTE", "RTE": "RTE", "JUNCTION": "JCT", "JCT": "JCT",
    "MOUNT": "MT", "MT": "MT", "MOUNTAIN": "MTN", "MTN": "MTN",
    "APARTMENT": "APT", "APT": "APT", "SUITE": "STE", "STE": "STE",
    "BUILDING": "BLDG", "BLDG": "BLDG", "FLOOR": "FL", "FL": "FL", "UNIT": "UNIT",
}

_UNIT_DESIGNATORS = {"APT", "STE", "UNIT", "BLDG", "FL"}
_DIRECTIONAL_TOKENS = {"N", "S", "E", "W", "NE", "NW", "SE", "SW"}
_HOUSE_UNIT_RE = re.compile(r"^(\d+)([A-Z])$")


def norm_street(v) -> str:
    """Canonicalizes a street address so common formatting/abbreviation
    differences don't look like different addresses (see split_street_unit(),
    street_compat())."""
    s = norm_text(v)
    if not s:
        return ""
    tokens = s.split(" ")
    unit_letter = ""
    m = _HOUSE_UNIT_RE.match(tokens[0])
    if m:
        tokens[0] = m.group(1)
        unit_letter = m.group(2)
    elif (len(tokens) > 1 and tokens[0].isdigit()
          and len(tokens[1]) == 1 and tokens[1] not in _DIRECTIONAL_TOKENS):
        unit_letter = tokens.pop(1)

    out = [_STREET_TOKEN_MAP.get(tok.rstrip("."), tok.rstrip(".")) for tok in tokens]

    if unit_letter and not any(t in _UNIT_DESIGNATORS for t in out):
        out += ["UNIT", unit_letter]

    return " ".join(out)


def norm_name(v) -> str:
    """Upper/trim; placeholder values ('[Unknown]', 'N/A', ...) become ''.

    PERFORMANCE: interns the result (sys.intern) when non-blank. First/Last/
    Middle/Suffix values repeat heavily in real person data (thousands of
    "SMITH"s, "JOHN"s, ...) - this is the string used to build every
    Name+Suffix bucket key (see bucket_candidate_pairs()), so making equal
    names share one string object lets tuple hashing/equality short-circuit
    on identity instead of a full character comparison every time. Interning
    a string never changes its value or equality behavior - purely a
    memory/speed optimization."""
    s = norm_text(v)
    core = re.sub(r"[^A-Z0-9]", "", s)
    if not core or core in NAME_PLACEHOLDERS:
        return ""
    return sys.intern(s)


SSN_MIN_KNOWN_OVERLAP = 4


def _ssn_pattern(v) -> str:
    """9-char digit/'X' pattern (redacted chars as 'X'), or a string of the
    wrong length if the cell isn't SSN-shaped at all - shared by norm_ssn()
    and classify_ssn_issue() (via build_records(), which computes this ONCE
    per row and derives both norm_ssn()'s and classify_ssn_issue()'s answer
    from it, instead of each re-doing this same regex/replace work from
    scratch on the same cell)."""
    if v is None:
        return ""
    s = _numeric_cell_to_str(v).upper().replace("*", "X").replace("#", "X").replace("?", "X")
    return re.sub(r"[^0-9X]", "", s)


def norm_ssn(v) -> str:
    """9-char pattern of digits/'X' (redacted), or '' if unusable."""
    return _ssn_from_pattern(_ssn_pattern(v))


def _ssn_from_pattern(kept: str) -> str:
    """norm_ssn()'s decision given an already-computed _ssn_pattern()
    result."""
    if len(kept) != 9:
        return ""
    if "X" not in kept:
        return kept
    known = sum(c != "X" for c in kept)
    return kept if known >= SSN_MIN_KNOWN_OVERLAP else ""


def classify_ssn_issue(v, pattern=None) -> str:
    """Short human-readable reason if a NON-BLANK SSN was rejected by
    norm_ssn() - '' if already blank or accepted. Pass pattern (an
    already-computed _ssn_pattern() result for this same cell) to skip
    recomputing it - see build_records()."""
    if v is None:
        return ""
    raw = str(v).strip()
    if not raw or raw.lower() in ("nan", "none", "null"):
        return ""
    return _classify_ssn_from_pattern(_ssn_pattern(v) if pattern is None else pattern)


def _classify_ssn_from_pattern(kept: str) -> str:
    """classify_ssn_issue()'s decision given an already-computed
    _ssn_pattern() result and a confirmed-non-blank raw value."""
    if len(kept) != 9:
        return ""
    if "X" not in kept:
        return ""
    known = sum(c != "X" for c in kept)
    if known < SSN_MIN_KNOWN_OVERLAP:
        return (f"Masked SSN with only {known} known digit(s) "
                f"(< {SSN_MIN_KNOWN_OVERLAP} required) - too little to trust")
    return ""


_EXCEL_SERIAL_EPOCH = pd.Timestamp("1899-12-30")
_EXCEL_SERIAL_RE = re.compile(r"\d{1,6}(\.\d+)?")


def norm_dob(v) -> str:
    """Parses a DOB cell into 'YYYYMMDD', or '' if unparseable/blank - also
    handles a raw Excel SERIAL date number (e.g. '20037') showing up as
    plain text instead of a formatted date, which some upstream export
    pipelines do even when the final file is a CSV (see hc_script_old.py's
    fuller explanation of the underlying quirk, originally observed via the
    pyxlsb .xlsb reader)."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    if _EXCEL_SERIAL_RE.fullmatch(s):
        serial = int(float(s))
        if 1 <= serial <= 60000:
            ts = _EXCEL_SERIAL_EPOCH + pd.Timedelta(days=serial)
            return ts.strftime("%Y%m%d")
    ts = pd.to_datetime(s, errors="coerce")
    return "" if pd.isna(ts) else ts.strftime("%Y%m%d")


# COL_DOB's own header says the expected format is MM/DD/YYYY - used as an
# EXPLICIT format for _vectorized_dob_fast_pass() below. Deliberately not a
# generic/inferred format - see that function's docstring for why.
_DOB_FAST_FORMAT = "%m/%d/%Y"


def _vectorized_dob_fast_pass(raw_series: pd.Series) -> list:
    """PERFORMANCE: a vectorized FIRST PASS over the whole DOB column,
    called ONCE from build_records() before its per-row loop. Resolves a
    cell ONLY when it unambiguously matches COL_DOB's own documented format
    ("Full Date of Birth (MM/DD/YYYY)"), via an EXPLICIT format string -
    never pandas' batch format INFERENCE. Returns a list the same length as
    raw_series: the 'YYYYMMDD' key for every cell this resolves, or None
    for every cell it doesn't (blank, an Excel serial number, a
    differently-formatted date, malformed/masked, ...) - those still fall
    through to norm_dob() one row at a time, exactly as before.

    WHY NOT JUST VECTORIZE norm_dob() OUTRIGHT: naive vectorization -
    pd.to_datetime(whole_column, errors="coerce") with NO explicit format -
    was tried and rejected. On a column with genuinely mixed date formats,
    pandas infers ONE dominant format for the whole batch and applies it to
    every row, silently returning NaT for any value that doesn't match it -
    even values a scalar, per-value pd.to_datetime() call parses correctly.
    Verified empirically (pandas 3.0.3): '1990-01-01', 'Jan 5, 1985',
    '1985-1-5', '01-05-1985', and '5-Jan-1985' all parse fine one at a
    time, but come back NaT when parsed together in a mixed-format test
    column. Silently turning a real DOB into blank is exactly the failure
    mode the rest of this script goes out of its way to avoid (a blank DOB
    never conflicts, so it can silently bridge two different people
    together during matching) - so this function instead only ever
    resolves a cell against ONE EXPLICIT, hardcoded format, which - unlike
    inference - parses identically whether done one value at a time or as
    a whole-column batch (verified against norm_dob()'s scalar result
    across a battery of unambiguous/ambiguous/invalid test values before
    being trusted here).

    IMPLEMENTATION NOTE: builds the result via an explicit numpy object
    array, not Series.where(parsed.notna(), None) - that was tried first
    and is BUGGY: pandas' .where() does not reliably write a literal Python
    None into an object-dtype Series at masked positions when the value
    already there is pandas' own NaN-like sentinel (which
    parsed.dt.strftime() produces for every NaT row) - it can silently keep
    that pre-existing NaN instead of substituting None. That's a
    correctness trap here specifically: float('nan') is truthy and
    NaN != NaN, so a row that should fall back to norm_dob() (checked via
    `is not None` in build_records()) would instead keep a bogus dob of
    nan, making every such row look like it has a real DOB that conflicts
    with every other blank-DOB row's nan - silently shredding the DOB
    conflict guard for every row this happened to. Explicit numpy
    boolean-mask assignment (below) has no such ambiguity."""
    parsed = pd.to_datetime(raw_series, format=_DOB_FAST_FORMAT, errors="coerce")
    ok = parsed.notna().to_numpy()
    out = np.empty(len(raw_series), dtype=object)
    out[:] = None
    if ok.any():
        out[ok] = parsed[ok].dt.strftime("%Y%m%d").to_numpy()
    return out.tolist()


def classify_dob_issue(v, dob_key=None) -> str:
    """Short human-readable reason if a NON-BLANK DOB failed to parse via
    norm_dob() - '' if already blank or parsed successfully. Pass dob_key
    (an already-computed norm_dob()/fast-pass result for this same cell) to
    skip re-parsing it a second time - see build_records()."""
    if v is None:
        return ""
    raw = str(v).strip()
    if not raw or raw.lower() in ("nan", "none", "null"):
        return ""
    key = norm_dob(v) if dob_key is None else dob_key
    return "" if key else "Unparseable DOB value - treated as blank"


def parse_id_tokens(v) -> frozenset:
    """Splits an ID cell (Employee ID / Driver's License / Passport / Tax ID
    / Phone) into individual normalized tokens - cells may already contain
    multiple semicolon-joined IDs from an earlier merge."""
    if v is None:
        return frozenset()
    return frozenset(norm_text(p) for p in _numeric_cell_to_str(v).split(";") if norm_text(p))


# ------------------------------------------------------------
# 3) Record type - __slots__ for fast attribute access at scale
# ------------------------------------------------------------
class Rec:
    __slots__ = ("idx", "first", "last", "mid", "suffix", "dob", "ssn",
                 "dl_ids", "passport_ids", "taxids",
                 "addr", "city", "state", "zip", "province")

    def __init__(self, idx):
        self.idx = idx


def build_records(df: pd.DataFrame):
    """df MUST already have a 0..n-1 RangeIndex (see main()). Returns
    (recs, ssn_review, dob_review) - same as hc_script_old.py, minus the
    Employee ID/Phone/Email token sets (no longer used for matching - those
    columns are pulled straight from the raw dataframe at output-build time
    instead, via SEMICONLON_COLS/OTHER_MERGE_COLS)."""
    recs = []
    ssn_review = []
    dob_review = []
    col_pos = {c: p for p, c in enumerate(df.columns)}
    values = df.values
    fi, la, mi, sf, do, da, ss, dl, pp, tx, ad, ci, st, zp, pv = (
        col_pos[COL_FIRST], col_pos[COL_LAST], col_pos[COL_MIDDLE], col_pos[COL_SUFFIX],
        col_pos[COL_DOB], col_pos[COL_DOB_ALT], col_pos[COL_SSN], col_pos[COL_DL],
        col_pos[COL_PASSPORT], col_pos[TAXID_COL], col_pos[COL_ADDR], col_pos[COL_CITY],
        col_pos[COL_STATE], col_pos[COL_ZIP], col_pos[COL_PROVINCE],
    )
    docid_pos = col_pos[COL_DOCID]

    # Vectorized first pass for the whole DOB column, and again for its
    # secondary/legacy source COL_DOB_ALT ("DOBs" - see the module
    # docstring) - see _vectorized_dob_fast_pass(). Resolves the common
    # case (a cell that cleanly matches the MM/DD/YYYY format) for every
    # row in one batched call each, so the per-row loop below only needs
    # to fall back to a scalar norm_dob() call for the rows this didn't
    # resolve.
    dob_fast = _vectorized_dob_fast_pass(df[COL_DOB])
    dob_alt_fast = _vectorized_dob_fast_pass(df[COL_DOB_ALT])

    for i in range(len(df)):
        row = values[i]
        r = Rec(i)
        r.first = norm_name(row[fi])
        r.last = norm_name(row[la])
        r.mid = norm_name(row[mi])
        r.suffix = norm_name(row[sf])

        fast_dob = dob_fast[i]
        primary_dob = fast_dob if fast_dob is not None else norm_dob(row[do])
        fast_dob_alt = dob_alt_fast[i]
        alt_dob = fast_dob_alt if fast_dob_alt is not None else norm_dob(row[da])
        # COL_DOB_ALT ("DOBs") feeds the SAME date evidence as COL_DOB, not
        # an independent field (see module docstring) - only used when the
        # primary column itself has nothing usable.
        r.dob = primary_dob if primary_dob else alt_dob

        ssn_pattern = _ssn_pattern(row[ss])
        r.ssn = _ssn_from_pattern(ssn_pattern)

        r.dl_ids = parse_id_tokens(row[dl])
        r.passport_ids = parse_id_tokens(row[pp])
        r.taxids = parse_id_tokens(row[tx])
        r.addr = norm_street(row[ad])
        r.city = norm_text(row[ci])
        r.state = norm_text(row[st])
        r.zip = norm_text(row[zp])
        r.province = norm_text(row[pv])
        recs.append(r)

        # Both review checks reuse the pattern/key already computed above
        # instead of re-parsing the same cell a second time (see
        # classify_ssn_issue()/classify_dob_issue()'s optional params).
        reason = classify_ssn_issue(row[ss], pattern=ssn_pattern)
        if reason:
            ssn_review.append({
                "DOCID": row[docid_pos],
                "First Name": row[fi], "Last Name": row[la],
                "Original SSN": row[ss],
                "Remarks": reason,
            })

        # Checked independently against EACH source column's own raw text
        # and own parsed key - a value that got silently treated as blank
        # in either COL_DOB or COL_DOB_ALT is worth surfacing, regardless
        # of whether the OTHER column happened to supply a usable r.dob.
        dob_reason = classify_dob_issue(row[do], dob_key=primary_dob)
        if dob_reason:
            dob_review.append({
                "DOCID": row[docid_pos],
                "First Name": row[fi], "Last Name": row[la],
                "Original DOB": row[do],
                "Remarks": dob_reason,
            })
        dob_alt_reason = classify_dob_issue(row[da], dob_key=alt_dob)
        if dob_alt_reason:
            dob_review.append({
                "DOCID": row[docid_pos],
                "First Name": row[fi], "Last Name": row[la],
                "Original DOB": row[da],
                "Remarks": f"[{COL_DOB_ALT}] {dob_alt_reason}",
            })
    return recs, ssn_review, dob_review


# ------------------------------------------------------------
# 4) Phase 2 pairwise matching rules - each pair tested here already shares
#    the EXACT same (First Name, Last Name, Suffix) key (see
#    bucket_candidate_pairs()), so unlike the old script, none of these
#    functions need to check Name/Suffix themselves - only the ID evidence
#    itself, plus the same conflict guards as before.
# ------------------------------------------------------------
def dob_conflict(r1: Rec, r2: Rec) -> bool:
    """True when BOTH rows have a usable DOB and it genuinely disagrees."""
    return bool(r1.dob) and bool(r2.dob) and r1.dob != r2.dob


def ssn_conflict(r1: Rec, r2: Rec) -> bool:
    """True when BOTH rows have a usable, known SSN and it genuinely
    disagrees."""
    return bool(r1.ssn) and bool(r2.ssn) and r1.ssn != r2.ssn


def name_prefix_compat(a: str, b: str) -> bool:
    """True if two (already normalized) values are compatible: blank on
    either side, exactly equal, or one is a PREFIX of the other."""
    if not a or not b or a == b:
        return True
    return a.startswith(b) or b.startswith(a)


def mid_conflict(r1: Rec, r2: Rec) -> bool:
    """True when Middle Name is present (non-blank, non-prefix-compatible)
    on both sides and genuinely disagrees. First/Last/Suffix are already
    guaranteed equal by the Name+Suffix partition these two rows came from
    (see bucket_candidate_pairs()), so only Middle Name still needs an
    active-conflict guard here."""
    return not name_prefix_compat(r1.mid, r2.mid)


def compatible(a: str, b: str) -> bool:
    """True if either side is blank, or both sides are equal."""
    return not a or not b or a == b


def ssn_match(r1: Rec, r2: Rec) -> bool:
    """Step 3.1 - SSN: both rows have the same usable SSN, unless a
    genuinely differing DOB or Middle Name blocks it."""
    if dob_conflict(r1, r2):
        return False
    if mid_conflict(r1, r2):
        return False
    return bool(r1.ssn) and bool(r2.ssn) and r1.ssn == r2.ssn


def ssn_dob_match(r1: Rec, r2: Rec) -> bool:
    """Step 3.2 - SSN + DOB: same DOB (Name+Suffix already guaranteed equal
    by the partition), as long as SSN doesn't conflict (blank on either/both
    sides is fine) and Middle Name doesn't actively disagree."""
    if ssn_conflict(r1, r2):
        return False
    if mid_conflict(r1, r2):
        return False
    return bool(r1.dob) and r1.dob == r2.dob


def _id_match(ids1: frozenset, ids2: frozenset, r1: Rec, r2: Rec) -> bool:
    """Shared logic for Steps 3.3-3.5 (Driver's License / Passport / Tax
    ID): rows share at least one common ID token AND SSN/DOB are each
    either matching or blank on both sides, AND Middle Name doesn't
    actively disagree."""
    if not (ids1 and ids2 and not ids1.isdisjoint(ids2)):
        return False
    if mid_conflict(r1, r2):
        return False
    return compatible(r1.ssn, r2.ssn) and compatible(r1.dob, r2.dob)


def dl_match(r1: Rec, r2: Rec) -> bool:
    """Step 3.3 - Driver's License (see _id_match())."""
    return _id_match(r1.dl_ids, r2.dl_ids, r1, r2)


def passport_match(r1: Rec, r2: Rec) -> bool:
    """Step 3.4 - Passport Number (see _id_match())."""
    return _id_match(r1.passport_ids, r2.passport_ids, r1, r2)


def taxid_match(r1: Rec, r2: Rec) -> bool:
    """Step 3.5 - Tax Identification (see _id_match() and the TAXID_COL
    assumption note in the module docstring)."""
    return _id_match(r1.taxids, r2.taxids, r1, r2)


def zip5(v: str) -> str:
    """First 5 digits of a ZIP code, ignoring hyphens/spaces/ZIP+4 suffix."""
    digits = re.sub(r"[^0-9]", "", v)
    return digits[:5] if len(digits) >= 5 else ""


def split_street_unit(s: str) -> tuple:
    """Splits a normalized street into (base, unit) - see hc_script_old.py's
    fuller explanation."""
    tokens = s.split(" ") if s else []
    for i, tok in enumerate(tokens):
        if tok in _UNIT_DESIGNATORS:
            return " ".join(tokens[:i]), " ".join(tokens[i + 1:])
        if tok.startswith("#") and len(tok) > 1:
            return " ".join(tokens[:i]), " ".join([tok[1:]] + tokens[i + 1:])
    return s, ""


def street_compat(a: str, b: str) -> bool:
    """True if two normalized streets are equal, blank on either side, or
    the same base street with a unit suffix present on only one side."""
    if not a or not b or a == b:
        return True
    base_a, unit_a = split_street_unit(a)
    base_b, unit_b = split_street_unit(b)
    if not base_a or base_a != base_b:
        return False
    return not (unit_a and unit_b and unit_a != unit_b)


# Strict priority order - see Phase 2 in the module docstring. A group
# formed at an earlier (lower-numbered) LEVEL is LOCKED once that level
# finishes, same locking discipline as hc_script_old.py's LEVEL_ORDER.
LEVEL_SSN      = 1   # Step 3.1
LEVEL_SSNDOB   = 2   # Step 3.2
LEVEL_DL       = 3   # Step 3.3
LEVEL_PASSPORT = 4   # Step 3.4
LEVEL_TAXID    = 5   # Step 3.5

LEVEL_ORDER = [LEVEL_SSN, LEVEL_SSNDOB, LEVEL_DL, LEVEL_PASSPORT, LEVEL_TAXID]

LEVEL_NAMES = {
    LEVEL_SSN: "SSN",
    LEVEL_SSNDOB: "SSN + DOB",
    LEVEL_DL: "Driver's License",
    LEVEL_PASSPORT: "Passport",
    LEVEL_TAXID: "Tax Identification",
}

LEVEL_MATCH_FUNCS = {
    LEVEL_SSN: ssn_match,
    LEVEL_SSNDOB: ssn_dob_match,
    LEVEL_DL: dl_match,
    LEVEL_PASSPORT: passport_match,
    LEVEL_TAXID: taxid_match,
}


# ------------------------------------------------------------
# 4b) Multiprocessing workers for the pairwise clustering step - identical
#     mechanism to hc_script_old.py, just operating over the new 5-level
#     LEVEL_MATCH_FUNCS map above.
# ------------------------------------------------------------
_worker_recs = None


def _init_worker(recs):
    global _worker_recs
    _worker_recs = recs


def _match_chunk(level_chunk):
    level, chunk = level_chunk
    match_func = LEVEL_MATCH_FUNCS[level]
    return [(a, b) for a, b in chunk if match_func(_worker_recs[a], _worker_recs[b])]


def _chunk_pairs(pairs_list, target_chunks=200, min_chunk_size=200):
    n = len(pairs_list)
    if n == 0:
        return []
    chunk_size = max(min_chunk_size, -(-n // target_chunks))
    return [pairs_list[i:i + chunk_size] for i in range(0, n, chunk_size)]


# ------------------------------------------------------------
# 5) Union-Find (disjoint set) for transitive clustering
# ------------------------------------------------------------
class UnionFind:
    def __init__(self, n):
        self.parent = list(range(n))

    def find(self, x):
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[max(ra, rb)] = min(ra, rb)


# ------------------------------------------------------------
# 6) Blocking - every bucket key is prefixed with the row's exact
#    (First Name, Last Name, Suffix) key, so two rows can NEVER be bucketed
#    together (and therefore never unioned) unless they already share that
#    exact Name+Suffix - this is what makes Phase 2's per-name partitioning
#    a hard boundary rather than something enforced by a separate loop.
# ------------------------------------------------------------
def describe_bucket_key(key) -> str:
    """PII-safe label for a bucket's grouping field - NEVER the actual
    Name/ID value (bucket keys are built directly from real PII/PHI - see
    bucket_candidate_pairs())."""
    level = key[0]
    return f"{LEVEL_NAMES[level]} (within same Name+Suffix)"


def bucket_candidate_pairs(recs, known_idxs):
    """Returns (level_pairs, biggest_buckets) - same shape/purpose as
    hc_script_old.py's version, but only over known_idxs (rows with a real
    First AND Last Name - see Phase 1), and every bucket key includes the
    row's exact (first, last, suffix) tuple so a candidate pair can only
    ever be produced between two rows sharing that exact Name+Suffix."""
    buckets = defaultdict(list)
    for i in known_idxs:
        r = recs[i]
        name_key = (r.first, r.last, r.suffix)
        if r.ssn:
            buckets[(LEVEL_SSN, name_key, r.ssn)].append(i)
        if r.dob:
            buckets[(LEVEL_SSNDOB, name_key, r.dob)].append(i)
        for tok in r.dl_ids:
            buckets[(LEVEL_DL, name_key, tok)].append(i)
        for tok in r.passport_ids:
            buckets[(LEVEL_PASSPORT, name_key, tok)].append(i)
        for tok in r.taxids:
            buckets[(LEVEL_TAXID, name_key, tok)].append(i)

    level_pair_lists = {lvl: [] for lvl in LEVEL_ORDER}
    total = len(buckets)
    biggest_buckets = []
    for n, (key, idxs) in enumerate(buckets.items(), 1):
        progress("Bucketing", n, total)
        if len(idxs) < 2:
            continue
        biggest_buckets.append((len(idxs), key))
        if len(biggest_buckets) > 5:
            biggest_buckets.sort(key=lambda t: -t[0])
            del biggest_buckets[5:]
        level = key[0]
        level_pair_lists[level].extend(itertools.combinations(sorted(idxs), 2))
    biggest_buckets.sort(key=lambda t: -t[0])

    level_pairs = {}
    for lvl, pairs in level_pair_lists.items():
        if not pairs:
            level_pairs[lvl] = []
            continue
        arr = np.array(pairs, dtype=np.int64)
        # Bit-pack each (a, b) pair into one int64 (a << 32 | b) and dedup
        # via a 1D np.unique instead of np.unique(..., axis=0) on the 2D
        # array - a 2D axis=0 unique pays for a row-wise structured-array
        # comparison/sort, while a 1D unique on plain int64 keys is a much
        # cheaper straight sort. Safe here because every pair already has
        # a < b (see itertools.combinations(sorted(idxs), 2) above), and no
        # row index can realistically approach 2**32 (a notification-list
        # CSV in the millions of rows is nowhere close), so the packing
        # never collides.
        packed = np.unique((arr[:, 0] << 32) | arr[:, 1])
        a = packed >> 32
        b = packed & 0xFFFFFFFF
        level_pairs[lvl] = list(zip(a.tolist(), b.tolist()))
    return level_pairs, biggest_buckets[:5]


# ------------------------------------------------------------
# 7) Merge helpers for building the output (unchanged from the prior
#    script - generic semicolon/date/address merge utilities)
# ------------------------------------------------------------
def semicolon_merge(values) -> str:
    """Distinct, non-blank values joined with '; ', first-seen order,
    original casing preserved; dedup key is upper/trimmed, split token-by-
    token on ';' first.

    PERFORMANCE: the overwhelming majority of calls (every un-merged/
    singleton output row, across every semicolon-merged column) pass a
    single value with no ';' in it - there is nothing to dedup against, so
    the tokenize/seen-set machinery below is pure overhead. That exact
    single-value/no-semicolon case is handled directly, byte-identically to
    what the general loop below would produce."""
    if len(values) == 1:
        v = values[0]
        if v is None:
            return ""
        s = str(v)
        if ";" not in s:
            raw = s.strip()
            return raw if norm_text(raw) else ""
    seen = set()
    out = []
    for v in values:
        if v is None:
            continue
        for tok in str(v).split(";"):
            raw = tok.strip()
            key = norm_text(raw)
            if not key or key in seen:
                continue
            seen.add(key)
            out.append(raw)
    return MERGE_SEP.join(out)


def dob_merge(values) -> str:
    """Like semicolon_merge(), but dedupes by the NORMALIZED date and always
    displays 'MM/DD/YYYY'. Same single-value/no-';' fast path as
    semicolon_merge() above, for the same reason."""
    if len(values) == 1:
        v = values[0]
        if v is None:
            return ""
        s = str(v)
        if ";" not in s:
            key = norm_dob(s.strip())
            return f"{key[4:6]}/{key[6:8]}/{key[0:4]}" if key else ""
    seen = set()
    out = []
    for v in values:
        if v is None:
            continue
        for tok in str(v).split(";"):
            raw = tok.strip()
            key = norm_dob(raw)
            if not key or key in seen:
                continue
            seen.add(key)
            out.append(f"{key[4:6]}/{key[6:8]}/{key[0:4]}")
    return MERGE_SEP.join(out)


DOCID_CHUNK_SIZE = 20_000


def split_docid_chunks(docid_str, max_chars=DOCID_CHUNK_SIZE):
    """Splits an already-merged DOCID string into chunks no longer than
    max_chars, breaking only at '; ' boundaries."""
    if len(docid_str) <= max_chars:
        return [docid_str]
    parts = docid_str.split(MERGE_SEP)
    chunks = []
    current = []
    current_len = 0
    for part in parts:
        added_len = len(part) + (len(MERGE_SEP) if current else 0)
        if current and current_len + added_len > max_chars:
            chunks.append(MERGE_SEP.join(current))
            current = [part]
            current_len = len(part)
        else:
            current.append(part)
            current_len += added_len
    if current:
        chunks.append(MERGE_SEP.join(current))
    return chunks


def fullest_value(raw_values, norm_values) -> str:
    """Longest raw value whose norm form is non-blank; '' if every value is
    blank/placeholder."""
    best, best_len = "", -1
    for raw, norm in zip(raw_values, norm_values):
        if not norm:
            continue
        raw = "" if raw is None else str(raw).strip()
        if len(raw) > best_len:
            best, best_len = raw, len(raw)
    return best


def has_variation(raw_values) -> bool:
    """True if 2+ distinct real raw values were seen for this field."""
    return len({norm_text(v) for v in raw_values if norm_text(v)}) > 1


def zip_key(v) -> str:
    z5 = zip5(norm_text(v))
    return z5 if z5 else norm_text(v)


def _address_field_norm(col, v) -> str:
    if col == COL_ZIP:
        return zip_key(v)
    if col == COL_ADDR:
        return norm_street(v)
    return norm_text(v)


def address_key(values) -> tuple:
    return tuple(_address_field_norm(col, v) for col, v in zip(ADDRESS_COLS, values))


def address_key_conflict(k1: tuple, k2: tuple) -> bool:
    """True if two normalized address keys genuinely disagree - blank on
    either side is never a conflict."""
    addr1, city1, state1, prov1, zip1, country1 = k1
    addr2, city2, state2, prov2, zip2, country2 = k2
    if not street_compat(addr1, addr2):
        return True
    pairs = ((city1, city2), (state1, state2), (prov1, prov2), (zip1, zip2), (country1, country2))
    return any(a and b and a != b for a, b in pairs)


def format_full_address(values) -> str:
    parts = []
    for v in values:
        s = "" if v is None else str(v).strip()
        if s and s.lower() not in ("nan", "none", "null"):
            parts.append(s)
    return ", ".join(parts)


def split_addresses(values_arr, addr_col_pos, group_idxs):
    """Returns (majority_values, other_address_string) for one merged group -
    Step 4.1: keeps the prior script's majority-address / "Other Address"
    clustering logic unchanged (see hc_script_old.py for the fuller
    explanation)."""
    key_order = []
    key_count = {}
    key_raw = {}
    for idx in group_idxs:
        row = values_arr[idx]
        raw = tuple(row[p] for p in addr_col_pos)
        key = address_key(raw)
        if key not in key_count:
            key_count[key] = 0
            key_raw[key] = raw
            key_order.append(key)
        key_count[key] += 1

    parent = list(range(len(key_order)))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[max(ra, rb)] = min(ra, rb)

    for i, j in itertools.combinations(range(len(key_order)), 2):
        if not any(key_order[i]) or not any(key_order[j]):
            continue
        if not address_key_conflict(key_order[i], key_order[j]):
            union(i, j)

    clusters = defaultdict(list)
    for i in range(len(key_order)):
        clusters[find(i)].append(i)

    def cluster_weight(positions):
        return sum(key_count[key_order[i]] for i in positions)

    def cluster_values(positions):
        out = []
        for col_i in range(len(ADDRESS_COLS)):
            raws = [key_raw[key_order[i]][col_i] for i in positions]
            norms = [key_order[i][col_i] for i in positions]
            out.append(fullest_value(raws, norms))
        return tuple(out)

    non_blank_clusters = [c for c in clusters.values()
                          if any(any(key_order[i]) for i in c)]
    candidates = non_blank_clusters or list(clusters.values())
    majority_cluster = max(candidates, key=lambda c: (cluster_weight(c), -min(c)))

    other_strings = []
    for c in clusters.values():
        if c is majority_cluster or not any(any(key_order[i]) for i in c):
            continue
        other_strings.append(format_full_address(cluster_values(c)))

    return cluster_values(majority_cluster), MERGE_SEP.join(other_strings)


# ------------------------------------------------------------
# 7b) Phase 1 - Unknown Name Split (Steps 1-2)
# ------------------------------------------------------------
def split_unknown_name_rows(recs):
    """Returns (known_idxs, unknown_idxs) - unknown_idxs is every row whose
    First Name AND/OR Last Name is blank/placeholder (Step 2: 'either First
    or Last name'). known_idxs is everything else - the only rows that ever
    enter Phase 2's matching cascade."""
    known_idxs, unknown_idxs = [], []
    for r in recs:
        (unknown_idxs if (not r.first or not r.last) else known_idxs).append(r.idx)
    return known_idxs, unknown_idxs


def build_unknown_entries(df: pd.DataFrame, recs, unknown_idxs) -> list:
    """Step 1/2: raw, UNMERGED rows for every unknown-name row - one output
    row per original input row, every original column preserved as-is,
    except First/Last Name is displayed as NAME_UNKNOWN on whichever side(s)
    were blank (the other side keeps its real value if it had one). An
    'Unknown Field' column records which side(s) triggered the split, and a
    'DOCID' passthrough column is included for cross-reference. This sheet
    is deliberately NOT deduped/merged against itself (see the module
    docstring's Phase 1 ASSUMPTION) - a name-less/partial-name row isn't
    trustworthy enough to identity-match on ID alone in this version."""
    out = []
    values_arr = df.values
    col_pos = {c: p for p, c in enumerate(df.columns)}
    first_pos, last_pos = col_pos[COL_FIRST], col_pos[COL_LAST]
    for i in unknown_idxs:
        row = values_arr[i]
        rec = recs[i]
        d = {c: row[p] for c, p in col_pos.items()}
        missing = []
        if not rec.first:
            d[COL_FIRST] = NAME_UNKNOWN
            missing.append("First Name")
        if not rec.last:
            d[COL_LAST] = NAME_UNKNOWN
            missing.append("Last Name")
        d["Unknown Field"] = " & ".join(missing)
        out.append(d)
    return out


# ------------------------------------------------------------
# 7c) Phase 3 - Blank-ID Fold (Step 5)
# ------------------------------------------------------------
ID_EVIDENCE_COLS = (COL_SSN, COL_DOB, COL_DL, COL_PASSPORT, TAXID_COL)


def _row_evidence_score(row: dict) -> int:
    """How many of the 5 ID fields (SSN/DOB/Driver's License/Passport/Tax
    ID) are populated on this already-built output row - 0-5. Used both to
    decide FILLED vs. BLANK-ID (score > 0) and, among 2+ FILLED candidates,
    to break a tie by richness (see fold_blank_id_rows())."""
    return sum(1 for c in ID_EVIDENCE_COLS if str(row.get(c, "")).strip())


def _merge_name_field(base_val: str, extra_val: str) -> str:
    """Fullest real value between two rows' same name field - falls back to
    base_val (never to '') if neither side has a real value, so folding a
    blank into a row whose own value is also blank doesn't erase anything."""
    merged = fullest_value([base_val, extra_val], [norm_name(base_val), norm_name(extra_val)])
    return merged or base_val


def _fold_row_into(base: dict, extra: dict) -> None:
    """Folds one output row ('extra') into another ('base'), in place -
    Step 4: DOCID/Employee ID/Phone/Email/Driver's License/Passport/Tax ID/
    every other OTHER_MERGE_COLS column, plus every address field and
    'Other Address', are all just semicolon_merge()'d as plain text (both
    sides already hold '; '-joined values from Phase 2's build, and neither
    side was confirmed to share a physical address with the other beyond
    the same Name+Suffix - appending rather than re-running the majority-
    address clustering is the safer choice for this cross-group fold)."""
    for col in [COL_DOCID] + OTHER_MERGE_COLS + ADDRESS_COLS + ["Other Address"]:
        base[col] = semicolon_merge([base.get(col, ""), extra.get(col, "")])
    base[COL_DOB] = dob_merge([base.get(COL_DOB, ""), extra.get(COL_DOB, "")])
    base[COL_FIRST] = _merge_name_field(base[COL_FIRST], extra[COL_FIRST])
    base[COL_MIDDLE] = _merge_name_field(base.get(COL_MIDDLE, ""), extra.get(COL_MIDDLE, ""))
    base[COL_LAST] = _merge_name_field(base[COL_LAST], extra[COL_LAST])
    base[COL_SUFFIX] = _merge_name_field(base.get(COL_SUFFIX, ""), extra.get(COL_SUFFIX, ""))
    base[COL_SSN] = fullest_value(
        [base.get(COL_SSN, ""), extra.get(COL_SSN, "")],
        [norm_ssn(base.get(COL_SSN, "")), norm_ssn(extra.get(COL_SSN, ""))])
    base["Rows Merged"] = base["Rows Merged"] + extra["Rows Merged"]
    base["Names Differ"] = base["Names Differ"] or extra["Names Differ"]
    base["Blank-ID Row Merged (Tie-Break)"] = (
        base.get("Blank-ID Row Merged (Tie-Break)", False)
        or extra.get("Blank-ID Row Merged (Tie-Break)", False))


def fold_blank_id_rows(known_rows: list) -> tuple:
    """Step 5 - POST-BUILD pass over Phase 2's built rows. Groups them by
    the exact (First Name, Last Name, Suffix) key (re-derived via
    norm_name() - every row in one Phase-2 group already shares this exact
    key, so recomputing it here reconstructs the same partition), then
    within any partition holding 2+ rows:
      - no FILLED row (none of the 5 ID fields populated on ANY row in the
        partition): fold every row into one - Name+Suffix alone is the only
        evidence there is, nothing is left unmerged.
      - exactly one FILLED row: fold every BLANK-ID row into it.
      - 2+ FILLED rows: for each BLANK-ID row, fold into whichever FILLED
        row has the richest evidence (_row_evidence_score()) IF there's a
        single richest one (flagged True in 'Blank-ID Row Merged (Tie-
        Break)'); if 2+ FILLED rows are equally rich, leave the BLANK-ID
        row standing alone and report it on the "Ambiguous Name-Group
        Review" sheet instead of guessing.

    Returns (final_rows, review_rows)."""
    groups = defaultdict(list)
    for i, row in enumerate(known_rows):
        key = (norm_name(row[COL_FIRST]), norm_name(row[COL_LAST]), norm_name(row.get(COL_SUFFIX, "")))
        groups[key].append(i)

    absorbed = set()
    review_rows = []

    for name_key, idxs in groups.items():
        if len(idxs) < 2:
            continue
        filled = [i for i in idxs if _row_evidence_score(known_rows[i]) > 0]
        blank = [i for i in idxs if _row_evidence_score(known_rows[i]) == 0]
        if not blank:
            continue   # nothing to fold in this partition

        if not filled:
            # No ID evidence anywhere in this Name+Suffix partition - fold
            # every row into the first (lowest-index) one.
            base_i, rest = blank[0], blank[1:]
            for i in rest:
                _fold_row_into(known_rows[base_i], known_rows[i])
                absorbed.add(i)
            continue

        if len(filled) == 1:
            base_i = filled[0]
            for i in blank:
                _fold_row_into(known_rows[base_i], known_rows[i])
                absorbed.add(i)
            continue

        # 2+ filled candidates - evidence tie-break per blank row.
        scores = {i: _row_evidence_score(known_rows[i]) for i in filled}
        best = max(scores.values())
        winners = [i for i in filled if scores[i] == best]
        for i in blank:
            if len(winners) == 1:
                _fold_row_into(known_rows[winners[0]], known_rows[i])
                known_rows[winners[0]]["Blank-ID Row Merged (Tie-Break)"] = True
                absorbed.add(i)
            else:
                review_rows.append({
                    "First Name": known_rows[i][COL_FIRST],
                    "Last Name": known_rows[i][COL_LAST],
                    "Suffix": known_rows[i].get(COL_SUFFIX, ""),
                    "DOCIDs": known_rows[i].get(COL_DOCID, ""),
                    "Remarks": (f"{len(winners)} equally-strong filled candidates share this "
                                f"Name+Suffix, tied on evidence ({best} of 5 ID fields each) - "
                                f"left unmerged, needs manual review"),
                })

    final_rows = [row for i, row in enumerate(known_rows) if i not in absorbed]
    return final_rows, review_rows


# ------------------------------------------------------------
# 8) Main
# ------------------------------------------------------------
def main() -> None:
    run_start = time.monotonic()
    print(f"Reading {INPUT_CSV} ...")
    t0 = time.monotonic()
    df = pd.read_csv(INPUT_CSV, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.reset_index(drop=True)

    missing = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing:
        raise SystemExit(
            f"These expected columns were not found in {INPUT_CSV}:\n"
            f"  {missing}\nColumns present:\n  {list(df.columns)}\n"
            "Fix the COL_*/OTHER_MERGE_COLS names in the CONFIG block."
        )
    print(f"  {len(df):,} rows read. ({time.monotonic() - t0:.1f}s)")

    t0 = time.monotonic()
    recs, ssn_review, dob_review = build_records(df)
    print(f"  Records built. ({time.monotonic() - t0:.1f}s)")
    if ssn_review:
        print(f"  {len(ssn_review):,} SSN value(s) were ignored as junk/unusable "
              f"- see the 'Junk SSN Review' sheet.")
    if dob_review:
        print(f"  {len(dob_review):,} DOB value(s) failed to parse and were "
              f"treated as blank - see the 'Junk DOB Review' sheet.")

    # ---- Phase 1: Unknown Name Split (Steps 1-2) ----
    known_idxs, unknown_idxs = split_unknown_name_rows(recs)
    print(f"  {len(unknown_idxs):,} row(s) have no usable First and/or Last "
          f"Name - moved to 'Unknown_Entries' (not merged). "
          f"{len(known_idxs):,} row(s) proceed to matching.")
    unknown_rows = build_unknown_entries(df, recs, unknown_idxs)

    # ---- Phase 2: Name+Suffix ID Cascade (Step 3) ----
    print("Clustering (blocked comparison, strict priority levels, scoped "
          "within exact Name+Suffix) ...")
    t0 = time.monotonic()
    level_pairs, biggest_buckets = bucket_candidate_pairs(recs, known_idxs)
    total_candidates = sum(len(p) for p in level_pairs.values())
    print(f"  {total_candidates:,} candidate pairs to test across "
          f"{len(LEVEL_ORDER)} priority levels. (bucketing took "
          f"{time.monotonic() - t0:.1f}s)")
    if biggest_buckets:
        print("  Largest buckets (candidate-pair hotspots - values withheld, "
              "PII/PHI):")
        for size, key in biggest_buckets:
            pair_count = size * (size - 1) // 2
            print(f"    {describe_bucket_key(key):<40} {size:>8,} rows "
                  f"-> {pair_count:>12,} pairs")

    uf = UnionFind(len(recs))
    locked = [False] * len(recs)
    group_ssn = [({r.ssn} if r.ssn else set()) for r in recs]
    group_dob = [({r.dob} if r.dob else set()) for r in recs]
    refused_ssn = 0
    refused_dob = 0
    refused_lock = 0
    root_size = [1] * len(recs)
    touched_roots = set()

    def try_union(a_idx, b_idx, level):
        nonlocal refused_ssn, refused_dob, refused_lock
        ra, rb = uf.find(a_idx), uf.find(b_idx)
        if ra == rb:
            return
        if locked[ra] and locked[rb]:
            refused_lock += 1
            return
        sa, sb = group_ssn[ra], group_ssn[rb]
        if sa and sb and sa.isdisjoint(sb):
            refused_ssn += 1
            return
        da, db = group_dob[ra], group_dob[rb]
        if da and db and da.isdisjoint(db):
            refused_dob += 1
            return
        uf.union(a_idx, b_idx)
        merged_root = min(ra, rb)
        group_ssn[merged_root] = sa | sb
        group_dob[merged_root] = da | db
        locked[merged_root] = locked[ra] or locked[rb]
        root_size[merged_root] = root_size[ra] + root_size[rb]
        touched_roots.add(merged_root)

    pool = None
    try:
        for level in LEVEL_ORDER:
            level_t0 = time.monotonic()
            pairs_list = list(level_pairs[level])
            total_pairs = len(pairs_list)
            label = f"Clustering L{level} ({LEVEL_NAMES[level]})"
            if total_pairs:
                match_func = LEVEL_MATCH_FUNCS[level]
                print(f"  Level {level} ({LEVEL_NAMES[level]}): "
                      f"{total_pairs:,} candidate pairs.")
                if total_pairs < PARALLEL_THRESHOLD:
                    for tested, (a_idx, b_idx) in enumerate(pairs_list, 1):
                        if match_func(recs[a_idx], recs[b_idx]):
                            try_union(a_idx, b_idx, level)
                        progress(label, tested, total_pairs)
                else:
                    if pool is None:
                        pool = Pool(processes=PARALLEL_WORKERS,
                                    initializer=_init_worker, initargs=(recs,))
                    chunks = _chunk_pairs(pairs_list)
                    done = 0
                    for matched in pool.imap_unordered(
                            _match_chunk, [(level, c) for c in chunks]):
                        for a_idx, b_idx in matched:
                            try_union(a_idx, b_idx, level)
                        done += 1
                        progress(label, done, len(chunks))

            for root in touched_roots:
                true_root = uf.find(root)
                if root_size[true_root] > 1:
                    locked[true_root] = True
            touched_roots.clear()

            if total_pairs:
                print(f"    Level {level} done. ({time.monotonic() - level_t0:.1f}s)")
    finally:
        if pool is not None:
            pool.close()
            pool.join()

    if refused_lock:
        print(f"  {refused_lock:,} candidate merge(s) were refused - would have "
              f"merged two groups already finalized by a higher-priority level.")

    groups = defaultdict(list)
    for i in known_idxs:
        groups[uf.find(i)].append(i)
    groups = list(groups.values())

    print(f"  {len(known_idxs):,} known-name rows -> {len(groups):,} groups "
          f"after Phase 2 ({len(known_idxs) - len(groups):,} rows collapsed by a match).")
    if refused_ssn:
        print(f"  {refused_ssn:,} candidate merge(s) were refused - would have "
              f"combined 2+ different real SSNs into one group.")
    if refused_dob:
        print(f"  {refused_dob:,} candidate merge(s) were refused - would have "
              f"combined 2+ different real DOBs into one group.")

    # ---- Step 4: build one row per Phase-2 group ----
    print("Building merged output (Step 4) ...")
    t0 = time.monotonic()
    SEMICOLON_COLS = [COL_DOCID] + OTHER_MERGE_COLS
    total_groups = len(groups)
    known_rows = []

    values_arr = df.values
    col_pos = {c: p for p, c in enumerate(df.columns)}
    semicol_col_pos = [col_pos[c] for c in SEMICOLON_COLS]
    addr_col_pos = [col_pos[c] for c in ADDRESS_COLS]
    dob_pos, dob_alt_pos, first_pos, last_pos, mid_pos, suffix_pos, ssn_pos, unique_id_pos = (
        col_pos[COL_DOB], col_pos[COL_DOB_ALT], col_pos[COL_FIRST], col_pos[COL_LAST],
        col_pos[COL_MIDDLE], col_pos[COL_SUFFIX], col_pos[COL_SSN], col_pos[COL_UNIQUE_ID],
    )

    for n, group_idxs in enumerate(groups, 1):
        progress("Building output", n, total_groups)

        if len(group_idxs) == 1:
            i = group_idxs[0]
            rv = values_arr[i]
            rec = recs[i]
            row = {c: semicolon_merge([rv[p]]) for c, p in zip(SEMICOLON_COLS, semicol_col_pos)}
            # COL_DOB_ALT ("DOBs") feeds the same date evidence as COL_DOB
            # and is never its own output column - see module docstring.
            row[COL_DOB] = dob_merge([rv[dob_pos], rv[dob_alt_pos]])
            # Reuse the already-normalized fields from build_records() (rec.
            # first/last/mid/suffix/ssn are byte-identical to norm_name()/
            # norm_ssn() on this same raw cell - just computed once already)
            # instead of re-normalizing the same value a second time here.
            row[COL_FIRST] = fullest_value([rv[first_pos]], [rec.first])
            row[COL_LAST] = fullest_value([rv[last_pos]], [rec.last])
            row[COL_MIDDLE] = fullest_value([rv[mid_pos]], [rec.mid])
            row[COL_SUFFIX] = fullest_value([rv[suffix_pos]], [rec.suffix])
            row[COL_SSN] = fullest_value([rv[ssn_pos]], [rec.ssn])
            # Unique_ID is the source DB's own row key, not a person-
            # identity signal - keep the (only) row's value as-is, never
            # semicolon-merged - see module docstring.
            row[COL_UNIQUE_ID] = rv[unique_id_pos]

            for c, p in zip(ADDRESS_COLS, addr_col_pos):
                raw = rv[p]
                norm = _address_field_norm(c, raw)
                row[c] = "" if not norm else ("" if raw is None else str(raw).strip())
            row["Other Address"] = ""

            row["Rows Merged"] = 1
            row["Names Differ"] = False
            row["Blank-ID Row Merged (Tie-Break)"] = False
            known_rows.append(row)
            continue

        sub_rows = [values_arr[i] for i in group_idxs]
        sub_recs = [recs[i] for i in group_idxs]

        row = {c: semicolon_merge([r[p] for r in sub_rows])
               for c, p in zip(SEMICOLON_COLS, semicol_col_pos)}
        # COL_DOB_ALT ("DOBs") feeds the same date evidence as COL_DOB
        # across every row in the group - see module docstring.
        row[COL_DOB] = dob_merge([r[dob_pos] for r in sub_rows] + [r[dob_alt_pos] for r in sub_rows])

        # Same reuse as the singleton path above - sub_recs' fields are
        # already the norm_name()/norm_ssn() result for these exact cells.
        first_vals = [r[first_pos] for r in sub_rows]
        last_vals = [r[last_pos] for r in sub_rows]
        row[COL_FIRST] = fullest_value(first_vals, [rec.first for rec in sub_recs])
        row[COL_LAST] = fullest_value(last_vals, [rec.last for rec in sub_recs])
        row[COL_MIDDLE] = fullest_value([r[mid_pos] for r in sub_rows],
                                         [rec.mid for rec in sub_recs])
        suffix_vals = [r[suffix_pos] for r in sub_rows]
        row[COL_SUFFIX] = fullest_value(suffix_vals, [rec.suffix for rec in sub_recs])
        ssn_vals = [r[ssn_pos] for r in sub_rows]
        row[COL_SSN] = fullest_value(ssn_vals, [rec.ssn for rec in sub_recs])
        # Unique_ID: keep the FIRST (topmost) original row's value only -
        # group_idxs (and therefore sub_rows) is already in ascending
        # original-row-order, since known_idxs/groups are built by a single
        # sequential scan - see module docstring/split_unknown_name_rows().
        row[COL_UNIQUE_ID] = sub_rows[0][unique_id_pos]

        majority_addr_values, other_address = split_addresses(values_arr, addr_col_pos, group_idxs)
        for c, v in zip(ADDRESS_COLS, majority_addr_values):
            row[c] = v
        row["Other Address"] = other_address

        row["Rows Merged"] = len(group_idxs)
        row["Names Differ"] = has_variation(first_vals) or has_variation(last_vals)
        row["Blank-ID Row Merged (Tie-Break)"] = False
        known_rows.append(row)

    print(f"  Output built. ({time.monotonic() - t0:.1f}s)")

    # ---- Phase 3: Blank-ID Fold (Step 5) ----
    print("Folding blank-ID rows into their Name+Suffix's filled row(s) "
          "(Step 5) ...")
    t0 = time.monotonic()
    known_rows, ambiguous_review = fold_blank_id_rows(known_rows)
    print(f"  {len(known_rows):,} row(s) remain after Phase 3. "
          f"({time.monotonic() - t0:.1f}s)")
    if ambiguous_review:
        print(f"  {len(ambiguous_review):,} blank-ID row(s) left unmerged - "
              f"tied on evidence between 2+ filled candidates sharing the "
              f"same Name+Suffix. See the 'Ambiguous Name-Group Review' sheet.")

    # ---- DOCID overflow split - runs once, on the FINAL row set, so a
    #      DOCID list gained via Phase 3 folding is chunked correctly. ----
    docid_overflow_groups = 0
    max_docid_cols = 1
    for row in known_rows:
        docid_chunks = split_docid_chunks(row[COL_DOCID])
        row[COL_DOCID] = docid_chunks[0]
        for extra_i, chunk in enumerate(docid_chunks[1:], start=2):
            row[f"{COL_DOCID} {extra_i}"] = chunk
        if len(docid_chunks) > 1:
            docid_overflow_groups += 1
            max_docid_cols = max(max_docid_cols, len(docid_chunks))

    df_out = pd.DataFrame(known_rows)

    docid_extra_cols = [f"{COL_DOCID} {i}" for i in range(2, max_docid_cols + 1)]
    input_order = list(df.columns)
    extra_cols = [c for c in df_out.columns if c not in input_order and c not in docid_extra_cols]
    new_order = []
    for c in input_order:
        if c not in df_out.columns:
            continue
        new_order.append(c)
        if c == COL_DOCID:
            new_order.extend(docid_extra_cols)
    new_order.extend(extra_cols)
    df_out = df_out[new_order]

    # Final safety-net check ONLY - catches two different groups that
    # collapsed to an identical row across every column.
    dup_mask = df_out.duplicated(keep="first")
    n_dupes = int(dup_mask.sum())
    if n_dupes:
        df_out = df_out[~dup_mask].reset_index(drop=True)

    df_out = df_out.sort_values(["Rows Merged"], ascending=False).reset_index(drop=True)
    if n_dupes:
        print(f"  {n_dupes:,} fully-duplicate output row(s) removed "
              f"(final check only - every column was identical to another row).")

    if docid_overflow_groups:
        print(f"  {docid_overflow_groups:,} group(s) had a DOCID list too long for one "
              f"cell (> {DOCID_CHUNK_SIZE:,} chars) - split across up to "
              f"{max_docid_cols} '{COL_DOCID}' columns.")

    n_multi = (df_out["Rows Merged"] > 1).sum()
    print(f"  {n_multi:,} merged groups combine 2+ original rows.")
    biggest = df_out["Rows Merged"].max()
    print(f"  Largest merged group: {biggest:,} rows.")
    df_large_groups = df_out[df_out["Rows Merged"] > 50].sort_values("Rows Merged", ascending=False)
    if len(df_large_groups):
        print(f"  WARNING: {len(df_large_groups):,} group(s) merged >50 rows - "
              f"usually a shared junk value (e.g. a fake SSN). See the "
              f"'Large Group Review' sheet before trusting the output.")

    df_ssn_review = pd.DataFrame(ssn_review)
    df_dob_review = pd.DataFrame(dob_review)
    df_unknown_entries = pd.DataFrame(unknown_rows)
    df_ambiguous_review = pd.DataFrame(ambiguous_review)

    print(f"Writing {OUTPUT_CSV} and its companion review CSVs ...")
    t0 = time.monotonic()
    written_paths = _write_outputs(OUTPUT_CSV, {
        "Merged Notification Data": df_out,
        "Unknown_Entries": df_unknown_entries,
        "Junk SSN Review": df_ssn_review,
        "Junk DOB Review": df_dob_review,
        "Large Group Review": df_large_groups,
        "Ambiguous Name-Group Review": df_ambiguous_review,
    })
    print(f"  Written. ({time.monotonic() - t0:.1f}s)")

    print(f"Done -> {len(written_paths)} file(s) written "
          f"(total runtime {_format_duration(time.monotonic() - run_start)}):")
    for p in written_paths:
        print(f"  {p}")
    print("Reminder: save the output only to the secured/authorized folder for "
          "this data - never a desktop or personal drive. It contains SSN, "
          "DOB, and other PII/PHI.")


def _sheet_csv_path(base_path: str, sheet: str) -> str:
    """Derives a companion CSV path for a given output 'sheet' from
    OUTPUT_CSV - e.g. 'output.csv' -> 'output - Junk SSN Review.csv'. The
    main "Merged Notification Data" sheet is written to base_path itself,
    unchanged - see _write_outputs()."""
    if sheet == "Merged Notification Data":
        return base_path
    root, ext = os.path.splitext(base_path)
    return f"{root} - {sheet}{ext or '.csv'}"


def _write_outputs(base_path: str, sheets: dict) -> list:
    """Writes every named DataFrame to its own CSV file (CSV has no
    multi-sheet concept, so each of the old workbook's sheets becomes a
    separate file - see _sheet_csv_path()). utf-8-sig encoding (a UTF-8
    BOM) so Excel - which this output is routinely opened in despite being
    plain CSV - renders non-ASCII characters correctly instead of
    mis-detecting the encoding. Returns the list of paths written, in the
    same order as `sheets`."""
    paths = []
    for sheet, df in sheets.items():
        path = _sheet_csv_path(base_path, sheet)
        df.to_csv(path, index=False, encoding="utf-8-sig")
        print(f"  {sheet}: {len(df):,} row(s) -> {path}")
        paths.append(path)
    return paths


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
