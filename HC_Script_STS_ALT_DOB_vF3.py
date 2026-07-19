"""
HC_Script_STS_ALT_DOB_vF3.py

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

  PHASE 2 - Name ID Cascade (Step 3 of the new spec)
    Every remaining row (real First AND Last Name) is first partitioned by
    an EXACT normalized (First Name, Last Name) key - two rows can ONLY
    ever end up in the same merged person if they share this exact
    2-field key. Middle Name and Suffix are NOT part of it (see "Middle
    Name / Suffix combining" below) and never block a match either -
    unlike the old script's Middle Name conflict guard, a differing (or
    one-sided blank) Middle Name/Suffix can never keep two rows apart
    anymore. (The exact-match-no-prefix-tolerance behavior for First/Last
    themselves is unchanged from the prior version - see the "Name
    grouping" note below.) WITHIN each such name partition, rows are
    further clustered by 6 strict priority levels, applied in order and
    LOCKED same as the old script's LEVEL_ORDER mechanism (a level's
    groups, once formed, can gain new rows from a later/weaker level but
    can never be bridged into another already-locked group by a later
    level):
        Level 1: SSN                          (Step 3.1)
        Level 2: SSN + DOB                    (Step 3.2)
        Level 3: Driver's License             (Step 3.3)
        Level 4: Passport                     (Step 3.4)
        Level 5: Tax Identification           (Step 3.5)
        Level 6: Government-Issued ID Number  (Step 3.6)
    Since every pair compared already shares the exact same First+Last
    Name, that's no longer part of any level's match test - each level
    only checks the ID evidence itself, plus the same conflict guards as
    before (a genuinely differing DOB blocks a same-SSN match; a
    genuinely differing SSN blocks a same-DOB match).

    ASSUMPTION - per explicit instruction, Middle Name and Suffix no
    longer influence matching AT ALL - not the partition key, not a
    conflict guard. Two rows for the same First+Last Name merge freely
    regardless of Middle Name/Suffix (a blank Middle Name and a filled
    one, or "Jr" vs "Sr", can both merge on a shared SSN). This is a real
    risk: two DIFFERENT people who happen to share an exact First+Last
    Name and some ID evidence will no longer be told apart by a
    genuinely differing Suffix or Middle Name the way either briefly
    could in an earlier version of this script.

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
    Identification Number"), a real, distinct column in the source data.
    "Government-Issued ID Number" (Step 3.6) matches on COL_GOVID, added
    as its own priority level right after Tax Identification - previously
    an APPEND-ONLY field only (see OTHER_MERGE_COLS); it still doubles as
    one, same as Driver's License/Passport/Tax ID already do.

    A genuinely different real DOB is NEVER allowed to end up in the same
    merged person as a shared SSN - dob_conflict() blocks Step 3.1 outright,
    and every other ID level (3.2-3.6) requires DOB to be matching-or-blank
    (compatible()) - a row with a conflicting real DOB always stays its own
    separate line item. (COL_DOB_ALT / "DOBs" - a formerly-used secondary
    DOB source column - has been removed from this version: it was a data
    error in the source database and is no longer read, matched on, or
    displayed anywhere. COL_DOB is the only DOB source now.)

    One source column is handled specially, outside the 6-level cascade:
      - COL_UNIQUE_ID ("Unique_ID") - the source database's own primary
        key for a row, confirmed NOT to be a person-identity signal (never
        used for matching). The output simply keeps whichever value was on
        the FIRST (topmost) original row in a merged group and drops every
        other row's value - unlike every OTHER_MERGE_COLS field, it is
        never semicolon-merged (see main()'s Step 4 build loop and
        _fold_row_into(), which both deliberately exclude it).

    Middle Name / Suffix combining (Step 3.2 of the updated spec) - since
    neither is part of the partition key anymore, a merged group can
    genuinely contain different Middle Name/Suffix values, so each is
    resolved differently once a group is built (see main()'s Step 4 and
    _fold_row_into()):
      - Middle Name: the fullest/most complete value seen (same
        fullest_value() logic as First/Last/SSN) - "Max" meaning longest,
        not statistically most frequent.
      - Suffix: the MOST FREQUENT value seen (by normalized comparison),
        ties broken by (1) the fullest raw representation of the tied
        value, then (2) which tied value was seen first - see
        _suffix_tally_from_raw()/_render_suffix_from_tally(). Frequency is
        tracked via a small internal tally carried on every row (never an
        output column) so it stays correct across every later merge stage
        (Step 4's build AND Step 5's non-PII consolidation below), not
        just within one stage.

    ASSUMPTION - Name grouping is an EXACT normalized First+Last Name
    match (no prefix/initial tolerance). If two rows for the same person
    use an initial vs. a full name ("J" vs "Jeffrey"), they will NOT be
    grouped together in this version, unlike the old script's
    name_prefix_compat().

  PHASE 3 - Non-PII Consolidation (Step 5 of the updated spec)
    Once Phase 2 is done, each First+Last Name partition may still hold
    2+ separate output rows (only when nothing bridged them at any of
    the 6 levels). These are split into:
      - FILLED rows: at least one of SSN/DOB/Driver's License/Passport/
        Tax ID/Government-Issued ID Number is populated.
      - BLANK-ID rows: none of those six PII fields are populated (only a
        name, and maybe non-PII attributes - Address/Employee ID/Phone/
        Email/DOCID/other OTHER_MERGE_COLS fields).
    Every BLANK-ID row in a partition is folded together into ONE combined
    row FIRST, regardless of how many FILLED rows exist or how they'll
    eventually be resolved - two blank-ID rows sharing a First+Last Name
    have, by definition, no PII evidence that could conflict, so there's
    no reason to leave several as separate stragglers. That single
    combined row is then folded (its Address/Employee ID/Phone/Email/
    DOCID/etc. appended in) into:
      - the partition's one FILLED row, if there is exactly one;
      - with 2+ FILLED rows, FIRST a filled candidate that shares a REAL
        non-PII attribute with the combined blank row - the EXACT same
        full Address, or an overlapping value in ANY OTHER_MERGE_COLS
        field (Email, Phone, Employee ID, Contact Information,
        Work-Related Information, Data Subject Type, ... - see
        _shares_nonpii_attribute()). If that resolves to EXACTLY one
        candidate, fold there directly - a real shared attribute, not a
        guess, so NOT flagged in "Blank-ID Row Merged (Tie-Break)".
      - only if that's inconclusive (no candidate shares anything, or 2+
        conflictingly do) does it fall back to "the top combined entry":
        the richest FILLED row by ID evidence (most of the 6 fields
        populated), then, if THAT ties too, the one representing the
        MAJORITY of original rows ("Rows Merged") - flagged True in
        "Blank-ID Row Merged (Tie-Break)" this time, since it WAS a
        guess. Only if even this fallback is fully tied (same evidence
        AND same size) is the combined blank row left standing alone and
        reported on the "Ambiguous Name-Group Review" sheet instead of
        guessed at - now a much rarer outcome than before, since the
        evidence/size fallback almost always produces a single winner;
      - nothing further, if the partition has NO filled row at all - the
        combined blank row IS the final row, First+Last Name alone being
        the only evidence there is.
    ASSUMPTION - per explicit instruction, the non-PII attribute sweep
    checks EVERY OTHER_MERGE_COLS field, not just Address/Email/Phone -
    the broadest reading of "Tags, etc.". This carries a real risk: a
    coincidental shared value in a generic, low-information field (e.g.
    the same "Data Subject Type") could link two unrelated people who
    happen to share a name - accepted per explicit instruction, but worth
    knowing about. The Address match specifically is an EXACT normalized
    match only (no blank-tolerant/unit-suffix fuzziness, unlike the
    majority-address clustering in split_addresses()) - per explicit
    instruction.

INPUT  : a CSV file (INPUT_CSV below) with the columns listed in
         EXPECTED_COLS below - every column read and written as plain text
         (dtype=str), so a leading zero in a ZIP/ID or a masked SSN like
         '123-45-XXXX' is never silently reinterpreted as a number.
OUTPUT : ONE workbook (OUTPUT_XLSB below), with the six sheets below each
         as their own tab - written as an intermediate .xlsx (no Python
         library can write .xlsb directly) then converted to .xlsb via
         Excel COM automation, see _write_workbook()/_convert_to_xlsb().
         If that conversion isn't possible (pywin32 not installed, or no
         local Excel install - Windows only), the .xlsx is kept and
         reported as the real output instead - see main().
         - "Merged Notification Data": ONE ROW PER CONFIRMED PERSON
           (known-name rows only - see Unknown_Entries below).
           - First Name, Middle Name, Last Name, and SSN: the single
             fullest/most complete value among the merged rows.
           - Suffix: the single MOST FREQUENT value among the merged rows
             (not fullest - see the Phase 2 "Middle Name / Suffix
             combining" note and _render_suffix_from_tally()).
           - DOB: every distinct real date seen, "; "-joined, formatted
             MM/DD/YYYY (see dob_merge()).
           - Driver's License, Passport Number, Tax Identification Number,
             Government-Issued ID Number, Employee ID: every distinct ID
             token seen, deduplicated and "; "-joined (cells may already
             contain multiple semicolon-joined tokens - see
             parse_id_tokens()).
           - Every OTHER column (DOCIDs, Phone, Email, etc.): every
             distinct value seen, "; "-joined. DOCIDs spills into
             "DOCIDs 2", "DOCIDs 3", ... columns past DOCID_CHUNK_SIZE
             characters (see split_docid_chunks()) - Excel truncates a
             cell past 32,767 characters, so this keeps every DOCID
             intact and visible instead.
           - Address fields: the majority address stays in the normal
             columns (gaps filled from another row's fuller copy of the
             SAME address); every other distinct address goes into "Other
             Address", semicolon-joined (see split_addresses()).
           - "Rows Merged": how many original input rows this became.
           - "Names Differ": True if 2+ raw (pre-normalization) First or
             Last Name spellings were seen in the group.
           - "Blank-ID Row Merged (Tie-Break)": True only for a row that
             absorbed a blank-ID row via the evidence/size FALLBACK (see
             Phase 3) - i.e. no shared non-PII attribute resolved it, so
             it was decided by which filled candidate was richest/biggest
             instead. Worth a quick double-check. Never True for a row
             absorbed via a real shared non-PII attribute (Address/Email/
             Phone/etc. match) - that's not a guess.
         - "Unknown_Entries": every raw row pulled out in Phase 1 (First
           and/or Last Name blank/placeholder) - unmerged, one row each,
           with the blank name side(s) displayed as "[Unknown]".
         - "Junk SSN Review": every row where a non-blank SSN value was
           ignored as unusable (a masked SSN with too few known digits).
         - "Junk DOB Review": every row where a non-blank DOB value failed
           to parse (so it was treated as blank instead of silently
           dropped).
         - "Large Group Review": every merged group with more than 50 rows.
         - "Ambiguous Name-Group Review": every blank-ID row (already
           combined with any sibling blank-ID rows sharing its First+Last
           Name - see Phase 3) that had 2+ filled-row candidates sharing
           no distinguishing non-PII attribute, still tied after both the
           evidence-richness AND majority-size fallback tie-breaks - left
           unmerged (standing alone in Merged Notification Data) rather
           than guessed at. Expected to be rare/usually empty now, since
           the non-PII match and the evidence/size fallback together
           resolve almost every case.

This script does not touch the input file. Save the output only to the
secured/authorized folder for this data (never a desktop) - it contains
SSN, DOB, and other PII/PHI.

Designed for large row counts (uses "blocking" - only compares rows that
already share an exact First+Last Name key AND an exact SSN, DOB, Driver's
License, Passport, Tax ID, or Government-Issued ID Number - instead of
comparing every row to every other row).

Install once:
    pip install pandas numpy xlsxwriter
    pip install pywin32   # optional - only needed to also emit a .xlsb
                           # output (Excel COM automation; Windows + a
                           # local Excel install required); without it,
                           # the intermediate .xlsx is kept as the output

Run:
    python HC_Script_STS_ALT_DOB_vF3.py
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
# PERFORMANCE: on Windows, multiprocessing uses 'spawn', not 'fork' - every
# worker process needs the full `recs` list (one Rec object per row)
# re-pickled and sent to it individually, since there's no copy-on-write
# memory sharing like POSIX fork gets. Profiling on a 265K-row file (3
# synthetic junk-SSN groups producing 6+ million candidate pairs) showed
# this pickling/process-startup cost alone at ~57s - while actually
# testing those SAME 6 million pairs serially (no worker processes at
# all) took just ~7s, since each match function (ssn_match(), etc.) is
# just a handful of cheap attribute comparisons. Going parallel was a net
# LOSS by roughly 8x for this workload. Raised well above any pair count a
# realistic file should ever produce in one level - multiprocessing still
# exists as a safety net for a genuinely pathological file (a single junk
# value shared by tens of thousands of rows in one level), where the
# pickling cost would finally be worth paying.
PARALLEL_THRESHOLD = 10_000_000

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
INPUT_CSV  = "sample_10K.csv"
# Single output workbook, one tab per sheet below (see _write_workbook()).
# Writing .xlsb directly isn't possible with any Python library (pandas/
# openpyxl/xlsxwriter/pyxlsb only ever WRITE .xlsx) - this writes a
# .xlsx first, then converts THAT to .xlsb via Excel COM automation (see
# _convert_to_xlsb() - requires pywin32 + a local Excel install, Windows
# only), deleting the intermediate .xlsx once the .xlsb exists. If either
# pywin32 or Excel itself isn't available, the .xlsx is kept and reported
# as the real output instead - see main().
OUTPUT_XLSB = "STS notification merge output.xlsb"

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
COL_WORKINFO = "Work-Related Information"

# Tax Identification match key (Step 3.5) - a real column in the source
# data (confirmed against the actual export's headers), not the
# Government-Issued ID Number workaround this used before that was known.
TAXID_COL = COL_TAXID

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
# regardless of which level (if any) matched. Employee ID, Phone, and
# Email are APPEND-ONLY here (see Phase 2 in the module docstring) - SSN,
# DOB, Driver's License, Passport, Tax ID, and Government-Issued ID Number
# ALSO double as match keys (see LEVEL_ORDER), on top of being listed here
# for the generic semicolon-merge treatment.
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
    COL_WORKINFO,
    "Family Information",
    "Financial Account Information",
    "Demographic Information",
    "Biometric Data",
    "PI Notes",
    "Access Credentials (Non-Financial Account)",
]

EXPECTED_COLS = (
    [COL_DOCID, COL_FIRST, COL_LAST, COL_MIDDLE, COL_SUFFIX, COL_DOB,
     COL_SSN, COL_UNIQUE_ID]
    + ADDRESS_COLS + OTHER_MERGE_COLS
)

MERGE_SEP = "; "
EXCEL_MAX_ROWS = 1_048_576   # a real constraint again now that output is .xlsx/.xlsb

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
    "SMITH"s, "JOHN"s, ...) - all four are the strings used to build every
    Name bucket key (see bucket_candidate_pairs()), so making equal names
    share one string object lets tuple hashing/equality short-circuit on
    identity instead of a full character comparison every time. Interning
    a string never changes its value or equality behavior
    - purely a memory/speed optimization."""
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

# Per explicit instruction: Excel's own blank/placeholder date cells land at
# or near its epoch (serial 0-366ish all resolve to late 1899/1900) when
# misread as a real date - no real notification subject has an actual birth
# year that early, so ANY parsed result landing at or before this year is
# treated as junk/blank rather than a real date, regardless of which branch
# (Excel-serial guess or ordinary text parsing) produced it. This is the
# same "1900 really means 0/NULL" fix applied uniformly everywhere a DOB
# cell gets parsed (norm_dob(), the vectorized fast pass, and the Junk DOB
# Review classification), not just the serial-number path.
DOB_MIN_YEAR = 1900


def norm_dob(v) -> str:
    """Parses a DOB cell into 'YYYYMMDD', or '' if unparseable/blank/
    implausible (see DOB_MIN_YEAR) - also handles a raw Excel SERIAL date
    number (e.g. '20037') showing up as plain text instead of a formatted
    date, which some upstream export pipelines do even when the final file
    is a CSV (see hc_script_old.py's fuller explanation of the underlying
    quirk, originally observed via the pyxlsb .xlsb reader)."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    if _EXCEL_SERIAL_RE.fullmatch(s):
        serial = int(float(s))
        if 1 <= serial <= 60000:
            ts = _EXCEL_SERIAL_EPOCH + pd.Timedelta(days=serial)
            return "" if ts.year <= DOB_MIN_YEAR else ts.strftime("%Y%m%d")
    ts = pd.to_datetime(s, errors="coerce")
    return "" if pd.isna(ts) or ts.year <= DOB_MIN_YEAR else ts.strftime("%Y%m%d")


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
    # DOB_MIN_YEAR floor applies here too (see norm_dob()) - a fast-pass hit
    # landing at/before it falls through to norm_dob() below, which rejects
    # it the same way, instead of the fast pass silently accepting it.
    ok = parsed.notna().to_numpy() & (parsed.dt.year > DOB_MIN_YEAR).to_numpy()
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
    if key:
        return ""
    # Re-parse only for a clearer audit message - distinguishes "genuinely
    # unparseable" from "parsed fine but landed at/before DOB_MIN_YEAR"
    # (Excel's own blank/placeholder date-serial range, see norm_dob()).
    # Both are equally blank for matching/display purposes either way.
    if _EXCEL_SERIAL_RE.fullmatch(raw):
        serial = int(float(raw))
        if 1 <= serial <= 60000:
            ts = _EXCEL_SERIAL_EPOCH + pd.Timedelta(days=serial)
            if ts.year <= DOB_MIN_YEAR:
                return (f"Parsed as an Excel date serial landing at/before "
                        f"{DOB_MIN_YEAR} ({ts.strftime('%m/%d/%Y')}) - a blank/"
                        f"placeholder date, not a real birthdate - treated as blank")
    else:
        ts = pd.to_datetime(raw, errors="coerce")
        if not pd.isna(ts) and ts.year <= DOB_MIN_YEAR:
            return (f"Parsed date falls at/before {DOB_MIN_YEAR} "
                     f"({ts.strftime('%m/%d/%Y')}) - implausible as a real "
                     f"birthdate - treated as blank")
    return "Unparseable DOB value - treated as blank"


def parse_id_tokens(v) -> frozenset:
    """Splits an ID cell (Employee ID / Driver's License / Passport / Tax ID
    / Government-Issued ID Number / Phone) into individual normalized
    tokens - cells may already contain multiple semicolon-joined IDs from
    an earlier merge."""
    if v is None:
        return frozenset()
    return frozenset(norm_text(p) for p in _numeric_cell_to_str(v).split(";") if norm_text(p))


# ------------------------------------------------------------
# 3) Record type - __slots__ for fast attribute access at scale
# ------------------------------------------------------------
class Rec:
    __slots__ = ("idx", "first", "last", "mid", "suffix", "dob", "ssn",
                 "dl_ids", "passport_ids", "taxids", "govid_ids",
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
    fi, la, mi, sf, do, ss, dl, pp, tx, gv, ad, ci, st, zp, pv = (
        col_pos[COL_FIRST], col_pos[COL_LAST], col_pos[COL_MIDDLE], col_pos[COL_SUFFIX],
        col_pos[COL_DOB], col_pos[COL_SSN], col_pos[COL_DL],
        col_pos[COL_PASSPORT], col_pos[TAXID_COL], col_pos[COL_GOVID], col_pos[COL_ADDR],
        col_pos[COL_CITY], col_pos[COL_STATE], col_pos[COL_ZIP], col_pos[COL_PROVINCE],
    )
    docid_pos = col_pos[COL_DOCID]

    # Vectorized first pass over the whole DOB column - see
    # _vectorized_dob_fast_pass(). Resolves the common case (a cell that
    # cleanly matches the MM/DD/YYYY format) for every row in one batched
    # call, so the per-row loop below only needs to fall back to a scalar
    # norm_dob() call for the rows this didn't resolve. (COL_DOB_ALT/"DOBs" -
    # a formerly-used secondary DOB source - has been removed: it was a
    # data error in the source database and is no longer read anywhere.)
    dob_fast = _vectorized_dob_fast_pass(df[COL_DOB])

    for i in range(len(df)):
        row = values[i]
        r = Rec(i)
        r.first = norm_name(row[fi])
        r.last = norm_name(row[la])
        r.mid = norm_name(row[mi])
        r.suffix = norm_name(row[sf])

        fast_dob = dob_fast[i]
        r.dob = fast_dob if fast_dob is not None else norm_dob(row[do])

        ssn_pattern = _ssn_pattern(row[ss])
        r.ssn = _ssn_from_pattern(ssn_pattern)

        r.dl_ids = parse_id_tokens(row[dl])
        r.passport_ids = parse_id_tokens(row[pp])
        r.taxids = parse_id_tokens(row[tx])
        r.govid_ids = parse_id_tokens(row[gv])
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

        dob_reason = classify_dob_issue(row[do], dob_key=r.dob)
        if dob_reason:
            dob_review.append({
                "DOCID": row[docid_pos],
                "First Name": row[fi], "Last Name": row[la],
                "Original DOB": row[do],
                "Remarks": dob_reason,
            })
    return recs, ssn_review, dob_review


# ------------------------------------------------------------
# 4) Phase 2 pairwise matching rules - each pair tested here already shares
#    the EXACT same (First Name, Middle Name, Last Name, Suffix) key (see
#    bucket_candidate_pairs()), so unlike the old script, none of these
#    functions need to check Name themselves at all - not even Middle Name
#    or Suffix (both are now baked directly into the partition key, so
#    equality is already guaranteed) - only the ID evidence itself, plus
#    the same DOB/SSN conflict guards as before.
# ------------------------------------------------------------
def dob_conflict(r1: Rec, r2: Rec) -> bool:
    """True when BOTH rows have a usable DOB and it genuinely disagrees."""
    return bool(r1.dob) and bool(r2.dob) and r1.dob != r2.dob


def ssn_conflict(r1: Rec, r2: Rec) -> bool:
    """True when BOTH rows have a usable, known SSN and it genuinely
    disagrees."""
    return bool(r1.ssn) and bool(r2.ssn) and r1.ssn != r2.ssn


def compatible(a: str, b: str) -> bool:
    """True if either side is blank, or both sides are equal."""
    return not a or not b or a == b


def ssn_full(r: Rec) -> bool:
    """True only when r.ssn is a COMPLETE, unmasked 9-digit SSN (no 'X'
    placeholder characters) - see norm_ssn()/SSN_MIN_KNOWN_OVERLAP. A
    masked/partial SSN (e.g. 'XXX-XX-1234') is real evidence for the SSN
    conflict guard (ssn_conflict()) and for the SSN+DOB/DL/Passport/TaxID/
    GovID compatibility checks, but per explicit instruction it must NEVER
    by itself be enough to CONFIRM a Level-1 SSN match - two different
    people who happen to share the same last few known digits (a real risk
    at scale, especially common names) must not be merged on that alone."""
    return bool(r.ssn) and "X" not in r.ssn


def ssn_match(r1: Rec, r2: Rec) -> bool:
    """Step 3.1 - SSN: both rows have the same FULL (unmasked) SSN, unless
    a genuinely differing DOB blocks it. See ssn_full() - a masked/partial
    SSN is never sufficient for this level, even if both sides match
    exactly."""
    if dob_conflict(r1, r2):
        return False
    return ssn_full(r1) and ssn_full(r2) and r1.ssn == r2.ssn


def ssn_dob_match(r1: Rec, r2: Rec) -> bool:
    """Step 3.2 - SSN + DOB: same DOB (First+Last Name already guaranteed
    equal by the partition), as long as SSN doesn't conflict (blank on
    either/both sides is fine)."""
    if ssn_conflict(r1, r2):
        return False
    return bool(r1.dob) and r1.dob == r2.dob


def _id_match(ids1: frozenset, ids2: frozenset, r1: Rec, r2: Rec) -> bool:
    """Shared logic for Steps 3.3-3.6 (Driver's License / Passport / Tax
    ID / Government-Issued ID Number): rows share at least one common ID
    token AND SSN/DOB are each either matching or blank on both sides."""
    if not (ids1 and ids2 and not ids1.isdisjoint(ids2)):
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


def govid_match(r1: Rec, r2: Rec) -> bool:
    """Step 3.6 - Government-Issued ID Number (see _id_match()). Added as
    its own priority level right after Tax Identification (Step 3.5) - a
    distinct, real column in the source data, previously only an
    APPEND-ONLY field (see OTHER_MERGE_COLS)."""
    return _id_match(r1.govid_ids, r2.govid_ids, r1, r2)


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
LEVEL_GOVID    = 6   # Step 3.6

LEVEL_ORDER = [LEVEL_SSN, LEVEL_SSNDOB, LEVEL_DL, LEVEL_PASSPORT, LEVEL_TAXID,
               LEVEL_GOVID]

LEVEL_NAMES = {
    LEVEL_SSN: "SSN",
    LEVEL_SSNDOB: "SSN + DOB",
    LEVEL_DL: "Driver's License",
    LEVEL_PASSPORT: "Passport",
    LEVEL_TAXID: "Tax Identification",
    LEVEL_GOVID: "Government-Issued ID Number",
}

LEVEL_MATCH_FUNCS = {
    LEVEL_SSN: ssn_match,
    LEVEL_SSNDOB: ssn_dob_match,
    LEVEL_DL: dl_match,
    LEVEL_PASSPORT: passport_match,
    LEVEL_TAXID: taxid_match,
    LEVEL_GOVID: govid_match,
}


# ------------------------------------------------------------
# 4b) Multiprocessing workers for the pairwise clustering step - identical
#     mechanism to hc_script_old.py, just operating over the new 6-level
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
#    (First Name, Last Name) key, so two rows can NEVER be bucketed
#    together (and therefore never unioned) unless they already share
#    that exact First+Last Name - this is what makes Phase 2's per-name
#    partitioning a hard boundary rather than something enforced by a
#    separate loop. Middle Name and Suffix are NOT part of this key (see
#    the module docstring) - they're combined as display attributes for
#    whatever cluster forms (Middle Name: fullest value; Suffix: most
#    frequent value), never used to decide who merges with whom.
# ------------------------------------------------------------
def describe_bucket_key(key) -> str:
    """PII-safe label for a bucket's grouping field - NEVER the actual
    Name/ID value (bucket keys are built directly from real PII/PHI - see
    bucket_candidate_pairs())."""
    level = key[0]
    return f"{LEVEL_NAMES[level]} (within same First+Last Name)"


def bucket_candidate_pairs(recs, known_idxs):
    """Returns (level_pairs, biggest_buckets) - same shape/purpose as
    hc_script_old.py's version, but only over known_idxs (rows with a real
    First AND Last Name - see Phase 1), and every bucket key includes the
    row's exact (first, last) tuple so a candidate pair can only ever be
    produced between two rows sharing that exact First+Last Name."""
    buckets = defaultdict(list)
    for i in known_idxs:
        r = recs[i]
        name_key = (r.first, r.last)
        if ssn_full(r):
            buckets[(LEVEL_SSN, name_key, r.ssn)].append(i)
        if r.dob:
            buckets[(LEVEL_SSNDOB, name_key, r.dob)].append(i)
        for tok in r.dl_ids:
            buckets[(LEVEL_DL, name_key, tok)].append(i)
        for tok in r.passport_ids:
            buckets[(LEVEL_PASSPORT, name_key, tok)].append(i)
        for tok in r.taxids:
            buckets[(LEVEL_TAXID, name_key, tok)].append(i)
        for tok in r.govid_ids:
            buckets[(LEVEL_GOVID, name_key, tok)].append(i)

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


def dob_merge_fast(raw_values, precomputed_keys) -> str:
    """Same output as dob_merge(raw_values), but for each raw value that
    does NOT itself contain ';', uses the matching entry in
    precomputed_keys (build_records()'s already-resolved r.dob for that
    exact cell - see Rec) instead of re-parsing it via
    norm_dob()/pd.to_datetime(). This is the overwhelming majority of
    cells, so it skips almost all of the re-parsing that made dob_merge()
    the single biggest hotspot in the whole pipeline when it was called
    once per output group in Step 4 (profiled at ~150s on a 265K-row run).

    Only a raw value that already holds multiple semicolon-joined dates
    (a cell from an earlier merge) falls back to full per-token norm_dob()
    parsing, identical to dob_merge() - precomputed_keys has no way to
    represent multiple dates for one cell, since Rec only ever resolves
    ONE date per column per row.

    raw_values and precomputed_keys must be the same length and in the
    same order (each precomputed_keys[i] is the resolved key for
    raw_values[i])."""
    seen = set()
    out = []
    for v, key in zip(raw_values, precomputed_keys):
        if v is None:
            continue
        s = str(v)
        if ";" not in s:
            if not key or key in seen:
                continue
            seen.add(key)
            out.append(f"{key[4:6]}/{key[6:8]}/{key[0:4]}")
            continue
        for tok in s.split(";"):
            raw = tok.strip()
            k = norm_dob(raw)
            if not k or k in seen:
                continue
            seen.add(k)
            out.append(f"{k[4:6]}/{k[6:8]}/{k[0:4]}")
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


# ------------------------------------------------------------
# 7a2) Suffix - "most frequent value" (Step 3.2 of the updated spec).
# Unlike First/Last/Middle/SSN (fullest_value() - one pass, no history
# needed), the WINNING Suffix can change as more rows fold in over
# multiple merge stages (Step 4's ID-cascade build, then Step 5's non-PII
# consolidation) - "Jr" might lead 3-to-2 after Step 4, then a fold in
# Step 5 could tip it to "Sr" 3-to-4. So every row carries a small tally
# (normalized Suffix -> [count, fullest raw seen]) as internal bookkeeping
# alongside its displayed COL_SUFFIX value, updated at every merge/fold
# and never included in the output (see _merged_output_column_order() -
# only explicitly-listed columns make it into the final CSV/workbook).
# ------------------------------------------------------------
_SUFFIX_TALLY_KEY = "_suffix_tally"


def _suffix_tally_from_raw(raw_values) -> dict:
    """Builds a fresh {normalized_suffix: [count, fullest_raw]} tally from
    a list of raw Suffix cell values - blank/placeholder values (see
    norm_name()) contribute nothing. Dict insertion order doubles as the
    'first seen' tie-break in _render_suffix_from_tally()."""
    tally = {}
    for raw in raw_values:
        norm = norm_name(raw)
        if not norm:
            continue
        raw_str = "" if raw is None else str(raw).strip()
        if norm not in tally:
            tally[norm] = [0, raw_str]
        tally[norm][0] += 1
        if len(raw_str) > len(tally[norm][1]):
            tally[norm][1] = raw_str
    return tally


def _merge_suffix_tally(tally_a: dict, tally_b: dict) -> dict:
    """Combines two Suffix tallies (counts add; the fullest raw
    representation seen for each normalized key is kept) - used whenever
    two already-built rows fold together, so frequency is tracked
    correctly across multiple merge stages instead of being lost the
    moment a single 'winning' display value is picked."""
    merged = {k: list(v) for k, v in tally_a.items()}
    for k, (count, raw) in tally_b.items():
        if k not in merged:
            merged[k] = [0, raw]
        merged[k][0] += count
        if len(raw) > len(merged[k][1]):
            merged[k][1] = raw
    return merged


def _render_suffix_from_tally(tally: dict) -> str:
    """The most frequent Suffix in a tally - ties broken by (1) the
    fullest/longest raw representation of the tied value, then (2) which
    tied value was seen first (dict insertion order) - both purely for
    determinism, since a genuine tie has no other principled winner.
    Returns '' if the tally is empty (no real Suffix seen anywhere)."""
    if not tally:
        return ""
    best_key = None
    best_count = -1
    best_len = -1
    for pos, (norm, (count, raw)) in enumerate(tally.items()):
        if (count, len(raw)) > (best_count, best_len):
            best_key, best_count, best_len = norm, count, len(raw)
    return tally[best_key][1]


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

    def try_union_pair(i, j):
        if not address_key_conflict(key_order[i], key_order[j]):
            union(i, j)

    # PERFORMANCE: bucket non-blank keys by BASE street (see
    # split_street_unit()) before testing pairs, instead of the naive
    # O(D^2) all-pairs loop (D = distinct address keys in this group) -
    # a real cost for a huge merged group with many genuinely different
    # addresses (e.g. rows sharing a junk SSN). street_compat() - the only
    # non-strict-equality check inside address_key_conflict() - can only
    # ever be True for two non-blank streets sharing the same base street;
    # every OTHER field (city/state/zip/country) is a plain blank-tolerant
    # equality check. So two keys with DIFFERENT non-blank base streets are
    # GUARANTEED to conflict regardless of every other field - skipping
    # those pairs is lossless (see test_split_addresses.py: 300/300
    # randomized cases matched the naive O(D^2) result exactly; 66x faster
    # on a 2000-row group with mostly-distinct addresses). A key with a
    # BLANK street is compatible with any base street (street_compat()
    # treats blank as a wildcard), so it's still tested against every
    # bucket, same as it would be in the naive version.
    non_blank_idxs = [i for i in range(len(key_order)) if any(key_order[i])]
    blank_street_idxs = []
    by_base_street = defaultdict(list)
    for i in non_blank_idxs:
        street = key_order[i][0]
        if not street:
            blank_street_idxs.append(i)
        else:
            by_base_street[split_street_unit(street)[0]].append(i)

    for i, j in itertools.combinations(blank_street_idxs, 2):
        try_union_pair(i, j)
    for bucket in by_base_street.values():
        for i, j in itertools.combinations(bucket, 2):
            try_union_pair(i, j)
        for i in blank_street_idxs:
            for j in bucket:
                try_union_pair(i, j)

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
ID_EVIDENCE_COLS = (COL_SSN, COL_DOB, COL_DL, COL_PASSPORT, TAXID_COL, COL_GOVID)


def _row_evidence_score(row: dict) -> int:
    """How many of the 6 ID fields (SSN/DOB/Driver's License/Passport/Tax
    ID/Government-Issued ID Number) are populated on this already-built
    output row - 0-6. Used both to decide FILLED vs. BLANK-ID (score > 0)
    and, among 2+ FILLED candidates, to break a tie by richness (see
    fold_blank_id_rows())."""
    return sum(1 for c in ID_EVIDENCE_COLS if str(row.get(c, "")).strip())


def _merge_name_field(base_val: str, extra_val: str) -> str:
    """Fullest real value between two rows' same name field - falls back to
    base_val (never to '') if neither side has a real value, so folding a
    blank into a row whose own value is also blank doesn't erase anything."""
    merged = fullest_value([base_val, extra_val], [norm_name(base_val), norm_name(extra_val)])
    return merged or base_val


def _merge_dob_display(base_dob: str, extra_dob: str) -> str:
    """Dedupes/re-joins two ALREADY-merged, already-formatted "MM/DD/YYYY;
    MM/DD/YYYY" style DOB display strings (exactly as dob_merge()/
    dob_merge_fast() produce - the only two producers of COL_DOB) without
    calling norm_dob()/pd.to_datetime() at all - the dedup key is just a
    plain string-slice rearrangement (YYYY+MM+DD) of an already-validated
    'MM/DD/YYYY' token, so there's nothing left to re-parse or validate.
    Used only here, during a Phase 3 fold - see _fold_row_into()."""
    seen = set()
    out = []
    for s in base_dob.split(MERGE_SEP) + extra_dob.split(MERGE_SEP):
        if not s:
            continue
        mm, dd, yyyy = s.split("/")
        key = yyyy + mm + dd
        if key in seen:
            continue
        seen.add(key)
        out.append(s)
    return MERGE_SEP.join(out)


def _merge_address_tuple(a: tuple, b: tuple) -> tuple:
    """Gap-fills two address tuples confirmed NOT to conflict (see
    address_key_conflict()) into one - same per-field fullest_value() logic
    as split_addresses()'s cluster_values(), just over two tuples instead of
    a whole cluster."""
    out = []
    for col, va, vb in zip(ADDRESS_COLS, a, b):
        out.append(fullest_value([va, vb], [_address_field_norm(col, va), _address_field_norm(col, vb)]))
    return tuple(out)


def _fold_addresses(base: dict, extra: dict) -> None:
    """Folds 'extra''s address into 'base', in place - a Phase 3 equivalent
    of split_addresses()'s majority-address / 'Other Address' clustering
    (Step 4.1), instead of a flat semicolon_merge() of the raw columns.

    BUG THIS REPLACES: a flat semicolon_merge() over ADDRESS_COLS smears two
    genuinely different addresses across every column ('100 Main St; 900
    Second Ave', 'Springfield; Reno', ...) instead of keeping one real
    address in the normal columns and routing the other to 'Other Address' -
    most visible exactly where Phase 3 fires (rows with NO other ID
    evidence, where address is often the only thing that could tell two
    people apart) - confirmed via a same-shape control case that merges
    through Phase 2's split_addresses() instead (shares a real ID, so
    Phase 2 handles it) and produces the correctly-split result.

    Each side's un-conflicting address is gap-filled together (same as
    split_addresses()); a genuine conflict keeps the side representing MORE
    original rows as the primary address (reusing 'Rows Merged' - already
    tracked on every built row - as the same majority-weight signal
    split_addresses() computes from scratch), demoting the other side's
    address into 'Other Address' - never dropped. Each side's own
    pre-existing 'Other Address' entries are carried over and deduped
    alongside it."""
    base_tup = tuple(base.get(c, "") for c in ADDRESS_COLS)
    extra_tup = tuple(extra.get(c, "") for c in ADDRESS_COLS)
    base_key = address_key(base_tup)
    extra_key = address_key(extra_tup)
    base_other = [s for s in base.get("Other Address", "").split(MERGE_SEP) if s]

    if not any(base_key):
        new_primary = extra_tup
    elif not any(extra_key):
        new_primary = base_tup
    elif not address_key_conflict(base_key, extra_key):
        new_primary = _merge_address_tuple(base_tup, extra_tup)
    else:
        # Genuine conflict - the side with more original rows behind it
        # wins the primary columns; the other is demoted to "Other Address"
        # rather than silently blended into the winning columns.
        if extra.get("Rows Merged", 1) > base.get("Rows Merged", 1):
            new_primary, demoted = extra_tup, base_tup
        else:
            new_primary, demoted = base_tup, extra_tup
        demoted_str = format_full_address(demoted)
        if demoted_str:
            base_other.append(demoted_str)

    extra_other = [s for s in extra.get("Other Address", "").split(MERGE_SEP) if s]
    seen = set()
    merged_other = []
    for s in base_other + extra_other:
        key = norm_text(s)
        if not key or key in seen:
            continue
        seen.add(key)
        merged_other.append(s)

    for col, v in zip(ADDRESS_COLS, new_primary):
        base[col] = v
    base["Other Address"] = MERGE_SEP.join(merged_other)


def _fold_row_into(base: dict, extra: dict) -> None:
    """Folds one output row ('extra') into another ('base'), in place -
    Step 4: DOCID/Employee ID/Phone/Email/Driver's License/Passport/Tax ID/
    every other OTHER_MERGE_COLS column are semicolon_merge()'d as plain
    text. Address fields are handled separately by _fold_addresses() -
    majority-clustered rather than flatly appended, since two genuinely
    different addresses must never blend into one unreadable string (see
    that function's docstring)."""
    for col in [COL_DOCID] + OTHER_MERGE_COLS:
        base[col] = semicolon_merge([base.get(col, ""), extra.get(col, "")])
    # MUST run before "Rows Merged" is updated below - _fold_addresses()
    # uses each side's CURRENT (pre-fold) "Rows Merged" as its majority-
    # weight signal.
    _fold_addresses(base, extra)
    base[COL_DOB] = _merge_dob_display(base.get(COL_DOB, ""), extra.get(COL_DOB, ""))
    base[COL_FIRST] = _merge_name_field(base[COL_FIRST], extra[COL_FIRST])
    base[COL_MIDDLE] = _merge_name_field(base.get(COL_MIDDLE, ""), extra.get(COL_MIDDLE, ""))
    base[COL_LAST] = _merge_name_field(base[COL_LAST], extra[COL_LAST])
    # Suffix: merge the underlying tallies (NOT fullest_value - see the
    # module docstring) so frequency stays correct across this fold too,
    # then re-render the (possibly now different) most-frequent winner.
    base[_SUFFIX_TALLY_KEY] = _merge_suffix_tally(
        base.get(_SUFFIX_TALLY_KEY, {}), extra.get(_SUFFIX_TALLY_KEY, {}))
    base[COL_SUFFIX] = _render_suffix_from_tally(base[_SUFFIX_TALLY_KEY])
    base[COL_SSN] = fullest_value(
        [base.get(COL_SSN, ""), extra.get(COL_SSN, "")],
        [norm_ssn(base.get(COL_SSN, "")), norm_ssn(extra.get(COL_SSN, ""))])
    base["Rows Merged"] = base["Rows Merged"] + extra["Rows Merged"]
    base["Names Differ"] = base["Names Differ"] or extra["Names Differ"]
    base["Blank-ID Row Merged (Tie-Break)"] = (
        base.get("Blank-ID Row Merged (Tie-Break)", False)
        or extra.get("Blank-ID Row Merged (Tie-Break)", False))


def _address_signature(row: dict):
    """EXACT-match address signature for a row already built by Step 4 -
    the normalized full-address tuple (see address_key()), or None if
    every ADDRESS_COLS field is blank. A blank address is never a real
    signal (see _shares_nonpii_attribute()) - two rows with nothing on
    file must never be treated as 'matching' on that nothing."""
    raw = tuple(row.get(c, "") for c in ADDRESS_COLS)
    key = address_key(raw)
    return key if any(key) else None


def _nonpii_tokens(row: dict) -> dict:
    """{field: frozenset(tokens)} for every OTHER_MERGE_COLS field on an
    already-built output row (values are already '; '-joined - see
    semicolon_merge()) - used by _shares_nonpii_attribute() for Step 5's
    common-attribute matching. A blank field naturally contributes an
    empty frozenset, which can never overlap with anything (see
    parse_id_tokens())."""
    return {c: parse_id_tokens(row.get(c, "")) for c in OTHER_MERGE_COLS}


def _shares_nonpii_attribute(blank_row: dict, filled_row: dict) -> bool:
    """Step 5's "common Address/Email/Phone/Tags" check: True if
    blank_row and filled_row share a REAL non-PII attribute - the EXACT
    same full Address (see _address_signature() - an exact normalized
    match only, no blank-tolerant/unit-suffix fuzziness), or an
    overlapping token in ANY OTHER_MERGE_COLS field (Email, Phone,
    Employee ID, Contact Information, Work-Related Information, Data
    Subject Type, ... - the broadest reading of "Tags, etc.", per
    explicit instruction). ASSUMPTION/RISK: checking every OTHER_MERGE_COLS
    field this broadly means a coincidental shared value in a generic,
    low-information field (e.g. the same "Data Subject Type") could link
    two unrelated people - accepted per explicit instruction, but worth
    knowing about. Never Address alone partially, never blank-vs-blank
    (both signature helpers treat blank as "no signal", not a wildcard)."""
    a1 = _address_signature(blank_row)
    if a1 is not None and a1 == _address_signature(filled_row):
        return True
    blank_tokens = _nonpii_tokens(blank_row)
    filled_tokens = _nonpii_tokens(filled_row)
    return any(blank_tokens[c] and filled_tokens[c]
               and not blank_tokens[c].isdisjoint(filled_tokens[c])
               for c in OTHER_MERGE_COLS)


def fold_blank_id_rows(known_rows: list) -> tuple:
    """Step 5 - POST-BUILD pass over Phase 2's built rows: non-PII
    consolidation. Groups them by the exact (First Name, Last Name) key
    (re-derived via norm_name() - every row in one Phase-2 group already
    shares this exact key, so recomputing it here reconstructs the same
    partition - Middle Name/Suffix are NOT part of it, see module
    docstring), then within any partition holding 2+ rows:
      - Every BLANK-ID row (none of the 6 ID fields populated) is folded
        together into ONE combined blank row FIRST, regardless of how many
        FILLED candidates exist or whether they'll turn out to match/tie -
        two blank-ID rows sharing a First+Last Name have, by definition,
        no ID evidence that could conflict between them, so there's no
        reason to ever leave them as separate stragglers.
      - no FILLED row anywhere in the partition: that combined blank row
        IS the final row - First+Last Name alone is the only evidence
        there is, nothing is left unmerged.
      - exactly one FILLED row: fold the combined blank row into it.
      - 2+ FILLED rows: FIRST, look for a filled candidate that shares a
        REAL non-PII attribute with the combined blank row - the exact
        same Address, or an overlapping Email/Phone/Employee ID/or any
        other OTHER_MERGE_COLS value (see _shares_nonpii_attribute()). If
        that resolves to EXACTLY one candidate, fold there directly (not
        a guess - a real shared attribute). Only if that's inconclusive
        (zero candidates share anything, OR 2+ conflictingly do) does it
        fall back to the OLD tie-break: richest ID evidence
        (_row_evidence_score(), 0-6 fields), then, if that ties too, the
        MAJORITY of original rows ('Rows Merged') - "the top combined
        entry" for this Name. A single winner from either path is flagged
        True in 'Blank-ID Row Merged (Tie-Break)' only when the FALLBACK
        path decided it (a real non-PII match isn't a tie-break). Only if
        the fallback is STILL tied on both evidence and size is the
        combined blank row left standing alone, reported on the
        "Ambiguous Name-Group Review" sheet instead of guessed at - a rare
        case now, since the fallback almost always produces a single "top"
        candidate.

    Returns (final_rows, review_rows)."""
    groups = defaultdict(list)
    for i, row in enumerate(known_rows):
        key = (norm_name(row[COL_FIRST]), norm_name(row[COL_LAST]))
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

        # Combine every blank-ID row in this partition into one FIRST - see
        # docstring above. Applies whether or not a filled row even exists.
        blank_base, blank_rest = blank[0], blank[1:]
        for i in blank_rest:
            _fold_row_into(known_rows[blank_base], known_rows[i])
            absorbed.add(i)

        if not filled:
            # No ID evidence anywhere in this First+Last Name partition -
            # nothing further to do, the combined blank row stands as-is.
            continue

        if len(filled) == 1:
            _fold_row_into(known_rows[filled[0]], known_rows[blank_base])
            absorbed.add(blank_base)
            continue

        # 2+ filled candidates - FIRST try a real non-PII match (Address/
        # Email/Phone/Tags) before falling back to a guess.
        nonpii_matches = [i for i in filled
                           if _shares_nonpii_attribute(known_rows[blank_base], known_rows[i])]
        if len(nonpii_matches) == 1:
            _fold_row_into(known_rows[nonpii_matches[0]], known_rows[blank_base])
            absorbed.add(blank_base)
            continue

        # Fallback ("the top combined entry"): tie-break stage 1 - evidence
        # richness.
        ev_scores = {i: _row_evidence_score(known_rows[i]) for i in filled}
        best_ev = max(ev_scores.values())
        ev_winners = [i for i in filled if ev_scores[i] == best_ev]

        # Tie-break stage 2 (only if stage 1 didn't decide it): prefer the
        # candidate with the most original rows already merged into it -
        # the majority cluster outweighs a same-richness straggler.
        if len(ev_winners) > 1:
            size_scores = {i: known_rows[i]["Rows Merged"] for i in ev_winners}
            best_size = max(size_scores.values())
            winners = [i for i in ev_winners if size_scores[i] == best_size]
        else:
            winners = ev_winners

        if len(winners) == 1:
            _fold_row_into(known_rows[winners[0]], known_rows[blank_base])
            known_rows[winners[0]]["Blank-ID Row Merged (Tie-Break)"] = True
            absorbed.add(blank_base)
        else:
            review_rows.append({
                "First Name": known_rows[blank_base][COL_FIRST],
                "Last Name": known_rows[blank_base][COL_LAST],
                "Suffix": known_rows[blank_base].get(COL_SUFFIX, ""),
                "DOCIDs": known_rows[blank_base].get(COL_DOCID, ""),
                "Remarks": (f"{len(winners)} equally-strong filled candidates share this "
                            f"First+Last Name, sharing no distinguishing non-PII attribute, "
                            f"tied on evidence ({best_ev} of 6 ID fields each) and on "
                            f"merged-row count ({size_scores[winners[0]]} rows each) - "
                            f"left unmerged, needs manual review"),
            })

    final_rows = [row for i, row in enumerate(known_rows) if i not in absorbed]
    return final_rows, review_rows


# ------------------------------------------------------------
# 7d) Final output column order - grouped by ROLE in the workflow rather
#     than mirroring the raw input file's column order, so a reviewer can
#     see at a glance what actually drove a match vs. what was just
#     appended along the way.
# ------------------------------------------------------------
# First/Last are the Phase 2 partition key; Middle/Suffix are combined
# display attributes only (Max/most-frequent - see module docstring), not
# part of the key - all four are still grouped together first in the
# output for readability.
_NAME_COLS = [COL_FIRST, COL_MIDDLE, COL_LAST, COL_SUFFIX]
# The 6-level merge cascade's OWN evidence fields, in the SAME priority
# order as Steps 3.1-3.6 (LEVEL_ORDER) - SSN first, Government-Issued ID
# Number last.
_MERGE_STEP_COLS = [COL_SSN, COL_DOB, COL_DL, COL_PASSPORT, TAXID_COL, COL_GOVID]
# Append-only contact/ID fields (Step 4.2, grouped with DOCID right after
# - see main()'s build loop) - never used to decide a match, only carried
# along once one is found some other way. Work-Related Information is
# grouped here per explicit instruction, alongside DOCID/Employee ID/
# Phone/Email, even though it's otherwise just another OTHER_MERGE_COLS
# field with no special handling of its own.
_CONTACT_ID_COLS = [COL_EMPID, COL_PHONE, COL_EMAIL, COL_WORKINFO]
# Meta/derived columns describing the merge itself, not source data.
_META_COLS = ["Rows Merged", "Names Differ", "Blank-ID Row Merged (Tie-Break)"]


def _merged_output_column_order(df_out_cols, docid_extra_cols) -> list:
    """Returns the column order for the 'Merged Notification Data' output:
      1. First/Middle/Last/Suffix (First+Last are the Phase 2 partition
         key; Middle/Suffix are combined display attributes only)
      2. The merge cascade's own evidence fields, in cascade order (SSN,
         DOB, Driver's License, Passport, Tax ID, Government-Issued ID
         Number)
      3. Address (majority address, then "Other Address") - not a match
         key anymore, but still identity-adjacent
      4. Append-only contact/ID fields (Employee ID, Phone, Email,
         Work-Related Information)
      5. DOCID, plus any "DOCID 2", "DOCID 3", ... overflow columns
      6. Every remaining append-only column (OTHER_MERGE_COLS, in their
         existing relative order) - the fields that never influence a
         match at all
      7. Unique_ID (first-row passthrough - see module docstring)
      8. Meta/derived columns (Rows Merged, Names Differ, Blank-ID Row
         Merged (Tie-Break))
    Only columns actually present in df_out_cols are included - this is
    just a preference ORDER, not a schema, so it's safe even if a column
    list changes later."""
    present = set(df_out_cols)
    placed = (set(_NAME_COLS) | set(_MERGE_STEP_COLS) | set(ADDRESS_COLS)
              | {"Other Address"} | set(_CONTACT_ID_COLS) | {COL_DOCID}
              | set(docid_extra_cols) | {COL_UNIQUE_ID} | set(_META_COLS))
    remaining_append_cols = [c for c in OTHER_MERGE_COLS if c not in placed and c in present]

    order = []
    order += [c for c in _NAME_COLS if c in present]
    order += [c for c in _MERGE_STEP_COLS if c in present]
    order += [c for c in ADDRESS_COLS if c in present]
    if "Other Address" in present:
        order.append("Other Address")
    order += [c for c in _CONTACT_ID_COLS if c in present]
    if COL_DOCID in present:
        order.append(COL_DOCID)
        order += [c for c in docid_extra_cols if c in present]
    order += remaining_append_cols
    if COL_UNIQUE_ID in present:
        order.append(COL_UNIQUE_ID)
    order += [c for c in _META_COLS if c in present]
    return order


def _split_groups_by_dob(groups: list, recs: list) -> list:
    """Defensive backstop run once, right after Phase 2's union-find groups
    are finalized: per explicit instruction, DOB must NEVER be displayed as
    a semicolon-joined list of 2+ genuinely different real dates for one
    merged person - if that happens, the rows behind it should never have
    been treated as one person, so they're split back into separate output
    rows (one per distinct DOB) instead of being merged.

    Every Phase 2 match level already guards against this directly
    (dob_conflict() blocks Level 1; every other level requires DOB to be
    matching-or-blank via compatible()), so a group reaching this function
    with 2+ distinct non-blank r.dob values should only happen from a
    data-quality edge case slipping past those guards (e.g. a junk numeric
    value that resolves to two different implausible dates on two
    different rows - see DOB_MIN_YEAR) - this function is the safety net
    for exactly that case, not the primary defense.

    Rows with NO usable DOB at all (r.dob == "") never conflict with
    anything, so they're folded into whichever real-DOB subgroup ends up
    LARGEST (majority) - same tie-break convention already used for
    addresses/suffixes elsewhere in this script - rather than left as their
    own orphan subgroup."""
    out = []
    split_count = 0
    for group_idxs in groups:
        if len(group_idxs) < 2:
            out.append(group_idxs)
            continue
        by_dob = defaultdict(list)
        for i in group_idxs:
            by_dob[recs[i].dob].append(i)
        real_dobs = [k for k in by_dob if k]
        if len(real_dobs) < 2:
            out.append(group_idxs)
            continue
        split_count += 1
        blank_idxs = by_dob.get("", [])
        if blank_idxs:
            majority_key = max(real_dobs, key=lambda k: len(by_dob[k]))
            by_dob[majority_key] = by_dob[majority_key] + blank_idxs
        for k in real_dobs:
            out.append(sorted(by_dob[k]))
    if split_count:
        print(f"  {split_count:,} group(s) had 2+ genuinely different real DOBs "
              f"despite matching on other evidence - split back into separate "
              f"rows per distinct DOB instead of merging (DOB is never "
              f"semicolon-joined across different real dates).")
    return out


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

    # ---- Phase 2: Name ID Cascade (Step 3) ----
    print("Clustering (blocked comparison, strict priority levels, scoped "
          "within exact First+Last Name) ...")
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

    # Safety net: DOB must never be shown as 2+ different real dates
    # blended into one merged row - see _split_groups_by_dob().
    groups = _split_groups_by_dob(groups, recs)

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
    dob_pos, first_pos, last_pos, mid_pos, suffix_pos, ssn_pos, unique_id_pos = (
        col_pos[COL_DOB], col_pos[COL_FIRST], col_pos[COL_LAST],
        col_pos[COL_MIDDLE], col_pos[COL_SUFFIX], col_pos[COL_SSN], col_pos[COL_UNIQUE_ID],
    )

    for n, group_idxs in enumerate(groups, 1):
        progress("Building output", n, total_groups)

        if len(group_idxs) == 1:
            i = group_idxs[0]
            rv = values_arr[i]
            rec = recs[i]
            row = {c: semicolon_merge([rv[p]]) for c, p in zip(SEMICOLON_COLS, semicol_col_pos)}
            row[COL_DOB] = dob_merge_fast([rv[dob_pos]], [rec.dob])
            # Reuse the already-normalized fields from build_records() (rec.
            # first/last/mid/suffix/ssn are byte-identical to norm_name()/
            # norm_ssn() on this same raw cell - just computed once already)
            # instead of re-normalizing the same value a second time here.
            row[COL_FIRST] = fullest_value([rv[first_pos]], [rec.first])
            row[COL_LAST] = fullest_value([rv[last_pos]], [rec.last])
            row[COL_MIDDLE] = fullest_value([rv[mid_pos]], [rec.mid])
            # Suffix: MOST FREQUENT value (not fullest) - see the module
            # docstring and _suffix_tally_from_raw()/_render_suffix_from_
            # tally(). The tally is internal bookkeeping only (never an
            # output column) so frequency stays correct across later
            # Step 5 folds too.
            row[_SUFFIX_TALLY_KEY] = _suffix_tally_from_raw([rv[suffix_pos]])
            row[COL_SUFFIX] = _render_suffix_from_tally(row[_SUFFIX_TALLY_KEY])
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
        # By the time a group reaches this point, _split_groups_by_dob() has
        # already guaranteed every row here shares the same (or a blank)
        # r.dob, so this should always collapse to at most one displayed
        # date - never a semicolon-joined list of genuinely different real
        # dates.
        row[COL_DOB] = dob_merge_fast([r[dob_pos] for r in sub_rows],
                                       [rec.dob for rec in sub_recs])

        # Same reuse as the singleton path above - sub_recs' fields are
        # already the norm_name()/norm_ssn() result for these exact cells.
        first_vals = [r[first_pos] for r in sub_rows]
        last_vals = [r[last_pos] for r in sub_rows]
        row[COL_FIRST] = fullest_value(first_vals, [rec.first for rec in sub_recs])
        row[COL_LAST] = fullest_value(last_vals, [rec.last for rec in sub_recs])
        row[COL_MIDDLE] = fullest_value([r[mid_pos] for r in sub_rows],
                                         [rec.mid for rec in sub_recs])
        # Suffix: MOST FREQUENT value across every row in this group, not
        # fullest - see the module docstring and _suffix_tally_from_raw().
        suffix_vals = [r[suffix_pos] for r in sub_rows]
        row[_SUFFIX_TALLY_KEY] = _suffix_tally_from_raw(suffix_vals)
        row[COL_SUFFIX] = _render_suffix_from_tally(row[_SUFFIX_TALLY_KEY])
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

    # ---- Phase 3: Non-PII Consolidation (Step 5) ----
    print("Consolidating blank-ID rows into their First+Last Name's "
          "filled row(s) via shared non-PII attributes, then evidence "
          "(Step 5) ...")
    t0 = time.monotonic()
    known_rows, ambiguous_review = fold_blank_id_rows(known_rows)
    print(f"  {len(known_rows):,} row(s) remain after Phase 3. "
          f"({time.monotonic() - t0:.1f}s)")
    if ambiguous_review:
        print(f"  {len(ambiguous_review):,} blank-ID row(s) left unmerged - "
              f"tied on evidence between 2+ filled candidates sharing the "
              f"same First+Last Name with no distinguishing non-PII "
              f"attribute. See the 'Ambiguous Name-Group Review' sheet.")

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

    if known_rows:
        df_out = pd.DataFrame(known_rows)
        docid_extra_cols = [f"{COL_DOCID} {i}" for i in range(2, max_docid_cols + 1)]
        df_out = df_out[_merged_output_column_order(df_out.columns, docid_extra_cols)]
    else:
        # No row survived at all (e.g. a genuinely empty input file, or
        # every row was an Unknown_Entries case) - pd.DataFrame([]) has NO
        # columns at all, which crashes every column-name reference below
        # (sort_values(["Rows Merged"]), etc.) - build a properly-headered,
        # zero-row frame instead so the workbook still writes cleanly with
        # the normal column layout, just no data rows.
        all_possible_cols = (_NAME_COLS + _MERGE_STEP_COLS + ADDRESS_COLS + ["Other Address"]
                              + _CONTACT_ID_COLS + [COL_DOCID] + OTHER_MERGE_COLS
                              + [COL_UNIQUE_ID] + _META_COLS)
        df_out = pd.DataFrame(columns=_merged_output_column_order(all_possible_cols, []))

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
    biggest = df_out["Rows Merged"].max() if len(df_out) else 0
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

    sheets = {
        "Merged Notification Data": df_out,
        "Unknown_Entries": df_unknown_entries,
        "Junk SSN Review": df_ssn_review,
        "Junk DOB Review": df_dob_review,
        "Large Group Review": df_large_groups,
        "Ambiguous Name-Group Review": df_ambiguous_review,
    }

    # Writing .xlsb directly isn't possible with any Python library - write
    # an intermediate .xlsx (multi-sheet, one tab per entry above) first,
    # then convert THAT to the real .xlsb output via Excel COM automation.
    temp_xlsx = os.path.splitext(OUTPUT_XLSB)[0] + ".xlsx"
    print(f"Writing {temp_xlsx} ...")
    t0 = time.monotonic()
    _write_workbook(temp_xlsx, sheets)
    print(f"  Written. ({time.monotonic() - t0:.1f}s)")

    print(f"Converting to {OUTPUT_XLSB} ...")
    t0 = time.monotonic()
    if _convert_to_xlsb(temp_xlsx, OUTPUT_XLSB):
        os.remove(temp_xlsx)
        print(f"  {OUTPUT_XLSB} written. ({time.monotonic() - t0:.1f}s)")
        print(f"Done -> {OUTPUT_XLSB} "
              f"({sum(len(df) for df in sheets.values()):,} rows across "
              f"{len(sheets)} tabs).")
    else:
        print(f"Done -> {temp_xlsx} "
              f"({sum(len(df) for df in sheets.values()):,} rows across "
              f"{len(sheets)} tabs) - see the message above for why the "
              f".xlsb conversion didn't run.")
    print("Reminder: save the output only to the secured/authorized folder for "
          "this data - never a desktop or personal drive. It contains SSN, "
          "DOB, and other PII/PHI.")
    print(f"Total run time: {_format_duration(time.monotonic() - run_start)} "
          f"(start to end).")


def _write_workbook(path: str, sheets: dict) -> None:
    """Writes every named DataFrame to its own tab in one .xlsx workbook,
    in the given order. A sheet longer than Excel's own row limit
    (EXCEL_MAX_ROWS) is spilled out to a companion CSV instead (with a
    one-line note left in its place in the workbook), since Excel would
    otherwise silently truncate it.

    REJECTED OPTIMIZATION - xlsxwriter's constant_memory=True option
    (engine_kwargs={"options": {"constant_memory": True}}) looked like a
    real win in isolation (~36% faster on an all-populated 300K-row
    sheet), but was found - via a minimal repro - to SILENTLY CORRUPT
    output whenever a DataFrame has scattered NaN/None values (entire
    rows can come back blank, with no error raised): pandas' to_excel()
    for the xlsxwriter engine does not guarantee writing cells in the
    strict row-sequential order constant_memory mode requires. Given this
    workbook always has real, scattered blank cells (unknown-name rows,
    unmatched ID fields, ...), that option must NOT be used here - a
    faster write is never worth silently dropped PII/PHI data."""
    with pd.ExcelWriter(path, engine="xlsxwriter") as xl:
        for sheet, df in sheets.items():
            if len(df) > EXCEL_MAX_ROWS:
                csv_name = f"{os.path.splitext(path)[0]} {sheet}.csv"
                df.to_csv(csv_name, index=False, encoding="utf-8-sig")
                print(f"  {sheet}: {len(df):,} rows exceed Excel's row limit "
                      f"-> {csv_name}")
                pd.DataFrame({"note": [f"{sheet} exported to {csv_name} (too large for one sheet)"]}
                             ).to_excel(xl, sheet_name=sheet, index=False)
            else:
                df.to_excel(xl, sheet_name=sheet, index=False)
                print(f"  {sheet}: {len(df):,} row(s)")


def _convert_to_xlsb(xlsx_path: str, xlsb_path: str) -> bool:
    """Saves a copy of the already-written .xlsx workbook as .xlsb via
    Excel COM automation (FileFormat 50 = xlsb) - the only reliable way to
    WRITE .xlsb, since no Python library (pandas/openpyxl/xlsxwriter/
    pyxlsb) can. Requires `pip install pywin32` and a local Excel install
    (Windows only); returns False (leaving the .xlsx as the real output)
    if either isn't available, rather than failing the whole run over an
    optional secondary format."""
    try:
        import win32com.client as win32
    except ImportError:
        print("  Skipping .xlsb conversion - pywin32 isn't installed "
              "(`pip install pywin32`). Kept the .xlsx output.")
        return False

    abs_xlsx, abs_xlsb = os.path.abspath(xlsx_path), os.path.abspath(xlsb_path)
    excel = None
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        # PERFORMANCE: this is a fresh, hidden Excel.Application instance
        # created just for this open/save round-trip (not the user's own
        # Excel, if they have one open) - turning off screen updates,
        # automatic recalculation, and event handlers is standard practice
        # for COM automation and meaningfully speeds up opening/saving a
        # large workbook. No need to restore these - the whole instance is
        # Quit() right after.
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.Calculation = -4135   # xlCalculationManual
        wb = excel.Workbooks.Open(abs_xlsx)
        try:
            wb.SaveAs(abs_xlsb, FileFormat=50)
        finally:
            wb.Close(SaveChanges=False)
        return True
    except Exception as exc:
        print(f"  Skipping .xlsb conversion - Excel COM automation failed: "
              f"{exc}. Kept the .xlsx output.")
        return False
    finally:
        if excel is not None:
            excel.Quit()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
