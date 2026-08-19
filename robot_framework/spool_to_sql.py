"""Parse a SAP spool tab-text export into rows.

Loading those rows into MSSQL lives in mssql_load.py. This module only reads the file
and works out which value belongs to which column.

WHY THIS IS NOT A PLAIN SPLIT ON TABS
-------------------------------------
The export is a printed classic ABAP list converted to text, not a data file. Two
consequences, both measured against a real 25,049-row export:

  1. Records wrap onto a second physical line, which carries its own sub-header
     (TilbF-Ref., TbF, TFB, User Name, Valoerdato).

  2. Data lines do not all carry the same number of tab fields. That export had six
     different field counts (20, 21, 65, 66, 67, 68) against a 68-field header, and the
     count does NOT tell you the alignment: rows of length 66 and 67 appear with both
     offsets. Reading fields straight off the header index therefore put the right value
     in the wrong column for 83.5% of rows - Periode, Bogfoeringsdato, Registr., Kl. and
     Aar were correct in only ~17% of them.

     Every row resolves to an offset of either 0 or 1, and the offset applies only from
     SHIFT_ANCHOR_COLUMN onwards; the leading columns (Bilagsnummer, BoL, PSP-element)
     are never shifted. Detecting the offset per row from two independent anchor columns
     put all nine spot-checked columns at 100% across all 25,049 rows.

So each row's offset is detected, not assumed, and a row whose offset cannot be
established raises rather than being silently mis-parsed.
"""

import re

# Columns used to detect a row's offset. Both must agree, which is what makes the
# detection trustworthy: a date alone could be matched by several columns, but a date
# with a 1-2 digit period immediately to its left is unambiguous in this layout.
ANCHOR_DATE_COLUMN = 'Bogføringsdato'
ANCHOR_PERIOD_COLUMN = 'Periode'

# Offsets to try, in order of preference.
CANDIDATE_OFFSETS = (0, 1)

_DATE = re.compile(r'\d{2}\.\d{2}\.\d{4}')
_PERIOD = re.compile(r'\d{1,2}')
_DIGITS = re.compile(r'\d+')

# Single-character flag columns on the continuation line. They are the cheap way to tell
# that a continuation line is misaligned - see _continuation_is_plausible.
SUB_FLAG_COLUMNS = ('TbF', 'TFB')
SUB_FLAG_MAX_LENGTH = 1


def parse_spool_file(file_path: str) -> tuple[list[str], list[dict], dict[str, int]]:
    """
    Parse a multi-section SAP spool tab-text export.

    Returns (column_names, rows, warnings):
      column_names - every column found, main line then continuation line
      rows         - one dict per record, keyed by the column name as it appears in the
                     header with SAP's right-alignment padding stripped
                     ('    Periode' -> 'Periode')
      warnings     - counts the caller should log; currently
                     'discarded_continuation_lines'
    """
    with open(file_path, encoding='cp1252', errors='replace') as handle:
        raw_lines = [line.rstrip('\r\n') for line in handle]

    warnings = {'discarded_continuation_lines': 0}

    header_indices = [i for i, line in enumerate(raw_lines) if _is_header_line(line)]
    if not header_indices:
        return [], [], warnings

    first = header_indices[0]
    main_headers = [h.strip() for h in raw_lines[first].split('\t')]
    sub_headers = ([h.strip() for h in raw_lines[first + 1].split('\t')]
                   if first + 1 < len(raw_lines) else [])

    main_col_map = _name_to_index(main_headers)
    sub_col_map = _name_to_index(sub_headers)

    if 'Bilagsnummer' not in main_col_map:
        raise ValueError("Column 'Bilagsnummer' not found in file header")
    for required in (ANCHOR_DATE_COLUMN, ANCHOR_PERIOD_COLUMN):
        if required not in main_col_map:
            raise ValueError(
                f"Anchor column {required!r} not found in file header, so a row's "
                f"column offset cannot be established. Header: {main_headers}"
            )

    bilag_idx = main_col_map['Bilagsnummer']
    date_idx = main_col_map[ANCHOR_DATE_COLUMN]
    period_idx = main_col_map[ANCHOR_PERIOD_COLUMN]
    shift_from = _shift_boundary(main_headers)

    all_col_names = list(main_col_map) + [k for k in sub_col_map if k not in main_col_map]

    rows: list[dict] = []
    index = 0
    while index < len(raw_lines):
        line = raw_lines[index]
        fields = line.split('\t')

        if _is_header_line(line):
            index += 2               # skip the header and its sub-header
            continue

        if _is_data_line(fields, bilag_idx):
            offset = _detect_offset(fields, date_idx, period_idx)
            if offset is None:
                raise ValueError(
                    f"Could not establish the column offset for the row on line "
                    f"{index + 1}: no candidate offset puts a date in "
                    f"{ANCHOR_DATE_COLUMN!r} and a period in {ANCHOR_PERIOD_COLUMN!r}. "
                    "The spool layout has probably changed."
                )

            row = {
                name: _field(fields, idx, offset, shift_from)
                for name, idx in main_col_map.items()
            }

            # The continuation line, if the next line is neither a new record nor a
            # header. It is read positionally - unlike the main line it has no usable
            # anchor, because Valoerdato is empty on 83% of rows and so cannot identify
            # an offset - and then sanity-checked before being trusted.
            if index + 1 < len(raw_lines):
                nxt = raw_lines[index + 1]
                nxt_fields = nxt.split('\t')
                if (nxt.strip()
                        and not _is_data_line(nxt_fields, bilag_idx)
                        and not _is_header_line(nxt)):
                    extra = {
                        name: (nxt_fields[idx].strip() if idx < len(nxt_fields) else '')
                        for name, idx in sub_col_map.items()
                        if name not in main_col_map
                    }
                    if _continuation_is_plausible(extra):
                        row.update(extra)
                    else:
                        warnings['discarded_continuation_lines'] += 1
                    index += 1

            for name in all_col_names:
                row.setdefault(name, '')

            rows.append(row)

        index += 1

    return all_col_names, rows, warnings


# ── internal helpers ──────────────────────────────────────────────────────────

def _is_header_line(line: str) -> bool:
    return 'Bilagsnummer' in [f.strip() for f in line.split('\t')]


def _is_data_line(fields: list[str], bilag_idx: int) -> bool:
    """A data line has a purely numeric document number in the Bilagsnummer slot.

    Bilagsnummer sits before the shifted region, so this test needs no offset.
    """
    if bilag_idx >= len(fields):
        return False
    return bool(_DIGITS.fullmatch(fields[bilag_idx].strip()))


def _name_to_index(headers: list[str]) -> dict[str, int]:
    """
    Map header name to field index, first occurrence winning.

    Duplicates are real: this layout repeats 'BoL' and 'OAr'. Keeping the first is
    deliberate - the first 'BoL' is the line item number the business key uses - but it
    does mean the later duplicate is unreachable. Widen this if one is ever needed.
    """
    mapping: dict[str, int] = {}
    for idx, name in enumerate(headers):
        if name:
            mapping.setdefault(name, idx)
    return mapping


def _shift_boundary(headers: list[str]) -> int:
    """
    First field index affected by a row's offset.

    The header carries a run of empty cells between the leading columns and the rest
    (indices 5-11 in the observed layout). Rows that omit one of those columns shift
    everything after the run, while the leading columns stay put. The boundary is
    therefore the first named column after the longest run of empty header cells.
    """
    longest_start = longest_len = 0
    run_start = run_len = 0
    for idx, name in enumerate(headers):
        if name:
            run_len = 0
            continue
        run_len = run_len + 1 if run_len else 1
        run_start = idx - run_len + 1
        if run_len > longest_len:
            longest_len, longest_start = run_len, run_start

    if longest_len == 0:
        return 0
    boundary = longest_start + longest_len
    while boundary < len(headers) and not headers[boundary]:
        boundary += 1
    return boundary


def _detect_offset(fields: list[str], date_idx: int, period_idx: int) -> int | None:
    """Return the offset under which both anchor columns hold the right shape."""
    for offset in CANDIDATE_OFFSETS:
        date_at = date_idx - offset
        period_at = period_idx - offset
        if not 0 <= date_at < len(fields) or not 0 <= period_at < len(fields):
            continue
        if (_DATE.fullmatch(fields[date_at].strip())
                and _PERIOD.fullmatch(fields[period_at].strip())):
            return offset
    return None


def _continuation_is_plausible(extra: dict[str, str]) -> bool:
    """
    Reject a continuation line whose single-character flag columns hold wide values.

    In the measured 25,049-row export, 3 rows had a shifted continuation line, which put
    25 characters of text into TFB - a CHAR(1) column in the target table. Loading that
    would fail the whole batch on a length check, so for those rows the continuation
    columns are dropped and counted instead. It costs six minor columns (TilbF-Ref.,
    TilbF-Org., TbF, TFB, User Name, Valoerdato) on 0.01% of rows; the main line, which
    carries the business key and every amount and date, is unaffected.
    """
    return all(len(extra.get(name, '')) <= SUB_FLAG_MAX_LENGTH
               for name in SUB_FLAG_COLUMNS)


def _field(fields: list[str], header_idx: int, offset: int, shift_from: int) -> str:
    """Read one column, applying the row's offset only past the shift boundary."""
    idx = header_idx if header_idx < shift_from else header_idx - offset
    return fields[idx].strip() if 0 <= idx < len(fields) else ''
