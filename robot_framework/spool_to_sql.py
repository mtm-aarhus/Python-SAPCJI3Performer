"""Parse a SAP spool tab-text export into rows.

Loading those rows into MSSQL lives in mssql_load.py. This module only reads
the file and works out the column layout from the header line.
"""

import re


def parse_spool_file(file_path: str) -> tuple[list[str], list[dict]]:
    """
    Parse a multi-section SAP spool tab-text export.

    Returns (column_names, rows) where each row is a dict keyed by the
    original column name found in the file header.  Column positions are
    read dynamically from the header line — nothing is hard-coded.
    """
    with open(file_path, encoding='cp1252', errors='replace') as f:
        raw_lines = [line.rstrip('\r\n') for line in f]

    # Locate every header line (a line that contains "Bilagsnummer" as a tab field)
    header_indices = [
        i for i, line in enumerate(raw_lines)
        if 'Bilagsnummer' in line.split('\t')
    ]
    if not header_indices:
        return [], []

    # All sections share the same column layout — derive structure from first occurrence
    first_hdr = header_indices[0]
    main_headers = raw_lines[first_hdr].split('\t')

    # The line immediately after the header contains extra sub-columns (TbF, TFB, …)
    sub_headers: list[str] = []
    if first_hdr + 1 < len(raw_lines):
        sub_headers = raw_lines[first_hdr + 1].split('\t')

    # Build name→index maps (first occurrence wins if a name appears twice)
    main_col_map: dict[str, int] = {}
    for idx, h in enumerate(main_headers):
        h = h.strip()
        if h:
            main_col_map.setdefault(h, idx)

    sub_col_map: dict[str, int] = {}
    for idx, h in enumerate(sub_headers):
        h = h.strip()
        if h:
            sub_col_map.setdefault(h, idx)

    bilag_idx = main_col_map.get('Bilagsnummer')
    if bilag_idx is None:
        raise ValueError("Column 'Bilagsnummer' not found in file header")

    # Combined ordered column list (sub-header columns that duplicate a main column are skipped)
    all_col_names = list(main_col_map) + [k for k in sub_col_map if k not in main_col_map]

    def _is_data_line(fields: list[str]) -> bool:
        val = fields[bilag_idx].strip() if bilag_idx < len(fields) else ''
        return bool(re.match(r'^\d+$', val))

    def _is_header_line(line: str) -> bool:
        return 'Bilagsnummer' in line.split('\t')

    rows: list[dict] = []
    i = 0
    while i < len(raw_lines):
        line = raw_lines[i]
        fields = line.split('\t')

        if _is_header_line(line):
            i += 2  # skip header + sub-header line
            continue

        if _is_data_line(fields):
            row: dict[str, str] = {
                col: fields[idx].strip() if idx < len(fields) else ''
                for col, idx in main_col_map.items()
            }

            # Peek at the next line: if it is non-blank, not a new data record,
            # and not a column header, treat it as the continuation line.
            if i + 1 < len(raw_lines):
                nxt = raw_lines[i + 1]
                nxt_fields = nxt.split('\t')
                if nxt.strip() and not _is_data_line(nxt_fields) and not _is_header_line(nxt):
                    for col, idx in sub_col_map.items():
                        if col not in row:
                            row[col] = nxt_fields[idx].strip() if idx < len(nxt_fields) else ''
                    i += 1  # consume the continuation line

            rows.append(row)

        i += 1

    return all_col_names, rows
