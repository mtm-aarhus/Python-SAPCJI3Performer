"""Parse SAP spool tab-text export and upsert rows into a SQLite database."""

import re
import sqlite3


def _sanitize(name: str) -> str:
    """Convert an arbitrary string to a safe SQL identifier."""
    return re.sub(r'[^a-zA-Z0-9_]', '_', name.strip()).strip('_') or 'col'


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


def upsert_to_sqlite(
    db_path: str,
    file_path: str,
    table_name: str = 'cji3_data',
    key_cols: tuple[str, ...] = ('Bilagsnummer', 'BoL'),
) -> int:
    """
    Parse *file_path* and upsert all rows into the SQLite database at *db_path*.

    *key_cols* is the composite primary key used to decide insert vs. overwrite.
    Change it here if the "position" column turns out to be something other than BoL.

    Returns the number of rows processed.
    """
    columns, rows = parse_spool_file(file_path)
    if not rows:
        return 0

    for k in key_cols:
        if k not in columns:
            raise ValueError(f"Key column '{k}' not found in file. Available: {columns}")

    conn = sqlite3.connect(db_path)
    try:
        _ensure_table(conn, table_name, columns, key_cols)
        _upsert_rows(conn, table_name, columns, rows)
        conn.commit()
    finally:
        conn.close()

    return len(rows)


# ── internal helpers ──────────────────────────────────────────────────────────

def _ensure_table(
    conn: sqlite3.Connection,
    table_name: str,
    columns: list[str],
    key_cols: tuple[str, ...],
) -> None:
    safe_table = _sanitize(table_name)
    safe_cols = [_sanitize(c) for c in columns]
    safe_keys = [_sanitize(k) for k in key_cols]

    col_defs = ', '.join(f'"{c}" TEXT' for c in safe_cols)
    pk_def = ', '.join(f'"{k}"' for k in safe_keys)
    conn.execute(
        f'CREATE TABLE IF NOT EXISTS "{safe_table}" ({col_defs}, PRIMARY KEY ({pk_def}))'
    )

    # Add columns that exist in the file but not yet in the table (schema evolution)
    existing_cols = {row[1] for row in conn.execute(f'PRAGMA table_info("{safe_table}")')}
    for col in safe_cols:
        if col not in existing_cols:
            conn.execute(f'ALTER TABLE "{safe_table}" ADD COLUMN "{col}" TEXT')


def _upsert_rows(
    conn: sqlite3.Connection,
    table_name: str,
    columns: list[str],
    rows: list[dict],
) -> None:
    safe_table = _sanitize(table_name)
    # Map original column name → sanitized SQL name
    col_pairs = [(orig, _sanitize(orig)) for orig in columns]
    safe_col_list = ', '.join(f'"{safe}" ' for _, safe in col_pairs)
    placeholders = ', '.join('?' for _ in col_pairs)

    sql = f'INSERT OR REPLACE INTO "{safe_table}" ({safe_col_list}) VALUES ({placeholders})'
    data = [tuple(row.get(orig, '') for orig, _ in col_pairs) for row in rows]
    conn.executemany(sql, data)
