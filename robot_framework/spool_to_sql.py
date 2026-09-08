"""Parse a SAP spool text export into rows.

Loading those rows into MSSQL lives in mssql_load.py. This module only reads the file
and works out which value belongs to which column.

THE FORMAT, AND WHY IT IS THIS ONE
----------------------------------
The robot exports via Spooljob > Videresend > 'Eksporter som tekst', which writes the
printed list as pipe-delimited, space-padded fixed-width text:

    |Bilagsnummer  |BoL|PSP-element        |OpV|beskrivelse
    |220144606     |  1|XA-1391100000-00020|  1|Sikring af bygning

Each record spans two physical lines: a long main line and a short continuation line
carrying TilbF-Ref., TilbF-Org., TbF, TFB, User Name and Valoerdato.

The obvious alternative, 'Tekst med tabulator', was tried first and does not work. It
places a tab at each print column boundary, so a cell that does not fill its column
changes the tab count and every later value on that row shifts. Measured against one
real 17,014-row export: seven different field counts across rows, and 12,872 rows
(75.7%) with a value that could not be converted - dates landing in Periode, quantities
in Aar, and so on. No per-row offset can repair it, because the drift accumulates as you
move along the row.

The same export read as pipe-delimited fixed-width: 61 fields on every row, and ZERO
bad values across nine typed columns and all 17,013 rows.

WHY SLICE BY OFFSET RATHER THAN SPLIT ON '|'
Three rows in that export contained a literal '|' inside a text value, which split()
turns into 63 fields instead of 61 and shifts the row. Slicing at the header's pipe
positions treats a stray pipe as just another character inside its column, and those
three rows parse correctly. Offsets are also what makes the format trustworthy in the
first place: they come from the printed layout, not from the data.
"""

import re

_DIGITS = re.compile(r'\d+')


def parse_spool_file(file_path: str) -> tuple[list[str], list[dict], dict[str, int]]:
    """
    Parse a SAP spool text export.

    Returns (column_names, rows, warnings):
      column_names - main-line columns then continuation-line columns
      rows         - one dict per record, keyed by the column name from the header
      warnings     - counts worth logging: 'rows_with_embedded_pipe',
                     'rows_without_continuation'
    """
    with open(file_path, encoding='cp1252', errors='replace') as handle:
        raw_lines = [line.rstrip('\r\n') for line in handle]

    warnings = {'rows_with_embedded_pipe': 0, 'rows_without_continuation': 0}

    header_indices = [i for i, line in enumerate(raw_lines) if _is_header_line(line)]
    if not header_indices:
        return [], [], warnings

    first = header_indices[0]
    main_ranges = _column_ranges(raw_lines[first])
    main_names = _slice_all(raw_lines[first], main_ranges)

    # The continuation sub-header is the line straight after the main header. Unlike the
    # main lines it does not start with a pipe, and it is SHORTER than its own data rows
    # - which is why the final column has to run to the end of the line being sliced
    # rather than to the end of the header.
    sub_ranges: list[tuple[int, int | None]] = []
    sub_names: list[str] = []
    if first + 1 < len(raw_lines):
        candidate = raw_lines[first + 1]
        if '|' in candidate and not _is_separator_line(candidate):
            sub_ranges = _column_ranges(candidate)
            sub_names = _slice_all(candidate, sub_ranges)

    # A genuine continuation line carries pipes at exactly the offsets the sub-header
    # defines. This is the test that keeps page furniture out: the export repeats a page
    # header and footer every 64 lines, and one of those footers follows a record's main
    # line, so a looser "next line that is not a record" test adopts it as the
    # continuation. That put 748 characters of footer into Valoerdato.
    sub_bars = set(_bar_positions(raw_lines[first + 1])) if sub_ranges else set()

    main_map = {name: i for i, name in enumerate(main_names) if name}
    sub_map = {name: i for i, name in enumerate(sub_names)
               if name and name not in main_map}

    if 'Bilagsnummer' not in main_map:
        raise ValueError(
            f"Column 'Bilagsnummer' not found in the header. Columns: {main_names}"
        )
    bilag_idx = main_map['Bilagsnummer']

    all_names = list(main_map) + list(sub_map)
    header_bars = _bar_positions(raw_lines[first])

    rows: list[dict] = []
    index = 0
    while index < len(raw_lines):
        line = raw_lines[index]

        if _is_header_line(line) or _is_separator_line(line):
            index += 1
            continue

        fields = _slice_all(line, main_ranges)
        if not _is_data_row(fields, bilag_idx):
            index += 1
            continue

        if _bar_positions(line) != header_bars:
            # A '|' inside a value. Harmless here - offset slicing ignores it - but
            # worth counting, since it would corrupt the row under split('|').
            warnings['rows_with_embedded_pipe'] += 1

        row = {name: fields[i] for name, i in main_map.items()}

        # Continuation line: the next line, when its pipes line up with the sub-header's.
        got_continuation = False
        if sub_ranges and index + 1 < len(raw_lines):
            nxt = raw_lines[index + 1]
            if (nxt.strip()
                    and sub_bars.issubset(_bar_positions(nxt))
                    and not _is_header_line(nxt)
                    and not _is_separator_line(nxt)
                    and not _is_data_row(_slice_all(nxt, main_ranges), bilag_idx)):
                extra = _slice_all(nxt, sub_ranges)
                for name, i in sub_map.items():
                    row[name] = extra[i] if i < len(extra) else ''
                got_continuation = True
                index += 1

        if not got_continuation:
            warnings['rows_without_continuation'] += 1

        for name in all_names:
            row.setdefault(name, '')

        rows.append(row)
        index += 1

    return all_names, rows, warnings


# ── internal helpers ──────────────────────────────────────────────────────────

def _bar_positions(line: str) -> list[int]:
    return [i for i, char in enumerate(line) if char == '|']


def _column_ranges(header_line: str) -> list[tuple[int, int | None]]:
    """
    Character range of each column, derived from the pipes in a header line.

    Main header lines both start and end with a pipe, so the columns are simply the
    gaps between consecutive pipes. The continuation sub-header starts with text, so
    the start of the line is a boundary too.

    The final range ends at None, meaning 'to the end of whatever line is sliced'. That
    is not cosmetic: the sub-header is 50 characters but its data rows are 52, so a
    range taken from the header would cut the last two characters off Valoerdato.
    """
    bars = _bar_positions(header_line)
    if not bars:
        return []

    edges = ([-1] if bars[0] != 0 else []) + bars
    ranges: list[tuple[int, int | None]] = [
        (a + 1, b) for a, b in zip(edges, edges[1:])
    ]
    last_start = edges[-1] + 1
    if last_start < len(header_line):
        ranges.append((last_start, None))
    elif ranges:
        ranges[-1] = (ranges[-1][0], None)
    return ranges


def _slice_all(line: str, ranges: list[tuple[int, int | None]]) -> list[str]:
    """Slice a line into its columns. A trailing pipe is not part of the last value."""
    out = []
    for start, end in ranges:
        piece = line[start:] if end is None else line[start:end]
        out.append(piece.strip().rstrip('|').strip() if end is None else piece.strip())
    return out


def _is_header_line(line: str) -> bool:
    return 'Bilagsnummer' in line


def _is_separator_line(line: str) -> bool:
    stripped = line.strip()
    return bool(stripped) and set(stripped) <= {'-', '|', '=', ' '} and '-' in stripped


def _is_data_row(fields: list[str], bilag_idx: int) -> bool:
    """A record's main line carries a purely numeric document number."""
    if bilag_idx >= len(fields):
        return False
    return bool(_DIGITS.fullmatch(fields[bilag_idx]))
