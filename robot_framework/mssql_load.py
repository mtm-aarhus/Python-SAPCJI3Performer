"""Load a parsed SAP CJI3 spool export into MSSQL (BI_Oekonomi.dbo.CJI3).

The spool headers are Danish and contain ae/oe/aa; the SQL columns are ASCII
transliterations. This mapping is the single point where the two meet - if SAP
changes a column heading, this dict is what needs updating, and an unknown
heading raises rather than being silently dropped.
"""

import uuid

import pyodbc

from robot_framework import config
from robot_framework.spool_to_sql import parse_spool_file


# SAP spool heading -> dbo.CJI3 column
COLUMN_MAP: dict[str, str] = {
    'Bilagsnummer'              : 'Bilagsnummer',
    'BoL'                       : 'BoL',
    'PSP-element'               : 'PSPElement',
    'beskrivelse'               : 'Beskrivelse',
    'Periode'                   : 'Periode',
    'Bogføringsdato'            : 'Bogfoeringsdato',
    'Medarbejders navn'         : 'MedarbejdersNavn',
    'Arbejdsplads tekst'        : 'ArbejdspladsTekst',
    'Modkontobetegnelse'        : 'Modkontobetegnelse',
    'Betegnelse'                : 'Betegnelse',
    'Se Faktura'                : 'SeFaktura',
    'Vrd./CO-områdevaluta'      : 'VaerdiCOOmraadevaluta',
    'ArbPlads'                  : 'ArbPlads',
    'Beskrivelse af modkontoen' : 'BeskrivelseAfModkontoen',
    'BilArt'                    : 'Bilagsart',
    'Originalobjektbet.'        : 'Originalobjektbetegnelse',
    'Bilagsdato'                : 'Bilagsdato',
    'Bilagstoptekst'            : 'Bilagstoptekst',
    'BME'                       : 'BME',
    'COVal'                     : 'COValuta',
    'FIL'                       : 'FIL',
    'Fulde navn'                : 'FuldeNavn',
    'Funktionsområde'           : 'Funktionsomraade',
    'Kapitalmidler'             : 'Kapitalmidler',
    'Modkonto'                  : 'Modkonto',
    'MA-nr.'                    : 'MANr',
    'ME'                        : 'ME',
    'Objekt'                    : 'Objekt',
    'OAr'                       : 'OAr',
    'Objektbetegnelse'          : 'Objektbetegnelse',
    'OArtbeskr.'                : 'OArtbeskrivelse',
    'OmkArtsbetegnelse'         : 'OmkArtsbetegnelse',
    'OmkostnArtsgrp.'           : 'OmkostnArtsgruppe',
    'Omk.art'                   : 'Omkostningsart',
    'Oper.'                     : 'Operation',
    'Originalobjekt'            : 'Originalobjekt',
    'Originalobjektart'         : 'Originalobjektart',
    'OrgOp'                     : 'OrgOp',
    'PartAA'                    : 'PartAA',
    'PFKo'                      : 'PFKo',
    'POmr'                      : 'POmr',
    'PartnKapMi'                : 'PartnKapMi',
    'Partnerobjekt'             : 'Partnerobjekt',
    'PAr'                       : 'PAr',
    'Partnerobjektart'          : 'Partnerobjektart',
    'Partnerobjektbetegnelse'   : 'Partnerobjektbetegnelse',
    'PObKl'                     : 'PObKl',
    'PartnerOS'                 : 'PartnerOS',
    'Projektdefinition'         : 'Projektdefinition',
    'RefBilNr'                  : 'RefBilagsnummer',
    'RT'                        : 'RT',
    'Registr.'                  : 'Registreringsdato',
    'Kl.'                       : 'Registreringstid',
    'År'                        : 'Aar',
    'Reg.mængde'                : 'Regmaengde',
    'Samlet mængde'             : 'SamletMaengde',
    'TilbF-Ref.'                : 'TilbFRef',
    'TilbF-Org.'                : 'TilbFOrg',
    'TbF'                       : 'TbF',
    'TFB'                       : 'TFB',
    'User Name'                 : 'UserName',
    'Valørdato'                 : 'Valoerdato',
}

# The business key. Position is not confirmed by Oekonomi yet, so it is fed from
# BoL for now - change this one line when they decide, nothing else moves.
POSITION_SOURCE = 'BoL'

STAGE_COLUMNS = ['Bilagsnummer', 'Position'] + [
    c for c in COLUMN_MAP.values() if c not in ('Bilagsnummer',)
]

# dbo.CJI3_Stage columns are NVARCHAR(255). Longer values are truncated
# rather than failing the whole batch; the count is logged so a shifted export
# does not pass unnoticed.
MAX_STAGE_LENGTH = 255


def connect(orchestrator_connection) -> pyodbc.Connection:
    """Open a connection to the BI_Oekonomi database."""
    sql_server = orchestrator_connection.get_constant(config.SQL_SERVER_CONSTANT).value
    conn_string = (
        f"DRIVER={{{config.SQL_DRIVER}}};"
        f"SERVER={sql_server};"
        f"DATABASE={config.SQL_DATABASE};"
        "Trusted_Connection=yes;"
    )
    return pyodbc.connect(conn_string)


def load_spool_file(
    orchestrator_connection,
    file_path: str,
    spool_job: str,
    udtraek_id: int | None,
) -> dict[str, int]:
    """
    Parse a spool export and merge it into dbo.CJI3.

    Rows are inserted as text into the staging table under a fresh BatchId and
    then merged by dbo.usp_CJI3_Merge, which does all date/number
    conversion, decides insert vs. update, writes the load log and - when
    udtraek_id is given - closes that window in the same transaction.

    Returns the proc's row counts.
    """
    _, rows = parse_spool_file(file_path)
    if not rows:
        raise ValueError(f"No data rows found in spool export: {file_path}")

    _verify_headers(rows)

    batch_id = str(uuid.uuid4())
    tuples = _to_stage_tuples(rows, batch_id)

    conn = connect(orchestrator_connection)
    try:
        cursor = conn.cursor()

        _verify_lengths(cursor, rows, orchestrator_connection)

        # Needs ODBC Driver 17/18; the legacy {SQL Server} driver does not support it.
        cursor.fast_executemany = True

        placeholders = ', '.join('?' for _ in range(len(STAGE_COLUMNS) + 1))
        columns = ', '.join(f'[{c}]' for c in ['BatchId'] + STAGE_COLUMNS)
        insert_sql = f"INSERT INTO dbo.CJI3_Stage ({columns}) VALUES ({placeholders})"

        for start in range(0, len(tuples), config.STAGE_CHUNK_SIZE):
            cursor.executemany(insert_sql, tuples[start:start + config.STAGE_CHUNK_SIZE])

        cursor.execute(
            "{CALL dbo.usp_CJI3_Merge (?, ?, ?, ?)}",
            batch_id, spool_job, file_path, udtraek_id,
        )
        counts = dict(zip([c[0] for c in cursor.description], cursor.fetchone()))

        conn.commit()
        return counts
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _verify_headers(rows: list[dict]) -> None:
    """
    Fail loudly if the spool contains a heading COLUMN_MAP does not know.

    A new or renamed SAP column would otherwise be dropped silently, and the
    data would look complete while quietly missing a field.
    """
    seen: set[str] = set()
    for row in rows:
        seen.update(row.keys())

    unknown = sorted(seen - set(COLUMN_MAP))
    if unknown:
        raise ValueError(
            f"Unknown spool column(s) {unknown}. Add them to COLUMN_MAP and to "
            "dbo.CJI3 / dbo.CJI3_Stage before loading."
        )


def _verify_lengths(cursor, rows: list[dict], orchestrator_connection) -> None:
    """
    Check every parsed value against the real column width before inserting anything.

    The widths are read from INFORMATION_SCHEMA rather than duplicated here, so this
    stays correct if a column is ever widened - there is one definition of the width
    and it lives in the database.

    Two reasons this exists. A value longer than its target column would otherwise
    surface as a bare "String or binary data would be truncated" from inside the
    MERGE, with no indication of which of 63 columns caused it. And silently cutting
    the value instead would put subtly wrong data into a finance table, which is worse
    than a failed run: the window stays IGang, gets reclaimed by the stale sweep, and
    retries - so a genuine SAP field that outgrew its column needs a human to widen it,
    which is the correct outcome.
    """
    cursor.execute(
        """
        SELECT COLUMN_NAME, CHARACTER_MAXIMUM_LENGTH
        FROM   INFORMATION_SCHEMA.COLUMNS
        WHERE  TABLE_SCHEMA = 'dbo'
          AND  TABLE_NAME   = 'CJI3'
          AND  CHARACTER_MAXIMUM_LENGTH > 0
        """
    )
    widths = {name: width for name, width in cursor.fetchall()}

    reverse = {sql: sap for sap, sql in COLUMN_MAP.items()}
    worst: dict[str, int] = {}

    for row in rows:
        for column, width in widths.items():
            source = POSITION_SOURCE if column == 'Position' else reverse.get(column)
            if source is None:
                continue
            length = len((row.get(source) or '').strip())
            if length > width and length > worst.get(column, 0):
                worst[column] = length

    if worst:
        detail = ', '.join(
            f"{column} needs {length}, column holds {widths[column]}"
            for column, length in sorted(worst.items())
        )
        raise ValueError(
            f"Value(s) too long for dbo.CJI3: {detail}. Widen the column in "
            "dbo.CJI3 and dbo.CJI3_Stage, or check whether the spool layout has "
            "shifted and put the wrong field in this column."
        )

    # Early warning well before anything fails, so a column can be widened during a
    # quiet moment rather than in response to a failed nightly run.
    tight = {
        column: width
        for column, width in widths.items()
        if _observed_max(rows, column, reverse) >= width * 0.8
    }
    if tight:
        orchestrator_connection.log_info(
            "Column(s) approaching their width limit: "
            + ', '.join(
                f"{column} at {_observed_max(rows, column, reverse)}/{width}"
                for column, width in sorted(tight.items())
            )
        )


def _observed_max(rows: list[dict], column: str, reverse: dict[str, str]) -> int:
    """Longest value seen for one target column across the parsed rows."""
    source = POSITION_SOURCE if column == 'Position' else reverse.get(column)
    if source is None:
        return 0
    return max((len((row.get(source) or '').strip()) for row in rows), default=0)


def _to_stage_tuples(rows: list[dict], batch_id: str) -> list[tuple]:
    """
    Turn parsed rows into staging tuples. Everything stays text.

    A value longer than the staging column is a hard error rather than a truncation:
    staging exists to carry the spool through unchanged, so quietly shortening a value
    here would defeat the point. In practice _verify_lengths fails first, since every
    target column is narrower than MAX_STAGE_LENGTH.
    """
    out: list[tuple] = []

    # SAP heading for each staging column, so lookups happen once per column.
    reverse = {sql: sap for sap, sql in COLUMN_MAP.items()}

    for row in rows:
        values: list[str | None] = [batch_id]
        for column in STAGE_COLUMNS:
            if column == 'Position':
                raw = row.get(POSITION_SOURCE, '')
            else:
                raw = row.get(reverse[column], '')

            raw = (raw or '').strip()
            if len(raw) > MAX_STAGE_LENGTH:
                raise ValueError(
                    f"Value for {column} is {len(raw)} chars, more than "
                    f"dbo.CJI3_Stage holds ({MAX_STAGE_LENGTH}). Widen the staging "
                    "column, or check whether the spool layout has shifted."
                )
            values.append(raw)
        out.append(tuple(values))

    return out
