"""Debuggable diagnostic: launch SAP if needed, dispatch a CJI3 extract, wait for the
spool job, and write an ANONYMISED copy of the raw export plus a layout report.

Press F5 in VS Code ("Debug: CJI3 layout sandbox"). Nothing here touches SQL and no
queue elements are created - it only drives SAP and writes files.

What it answers, which the old EXAMPLE.TXT cannot:
  * the exact raw column headings of the CURRENT layout, verbatim
  * how wide values in each column actually get, so create_CJI3.sql can be sized
    from fact instead of inference
  * which CJI3 date field %%DYN002 really is - it reads the on-screen label
  * whether the column layout is positionally reliable across rows

Anonymisation: every letter becomes 'a'/'A' and every digit '9', in place. Lengths,
punctuation, tabs and case pattern survive; content does not. There is no key and no
cipher, so nothing can be recovered. A date still reads '99.99.9999' and an amount
'9.999,99', which is all that is needed to judge layout. Column HEADINGS are kept
verbatim - they are metadata, not data.

Prerequisites: the two OpenOrchestrator environment variables below. SAP does NOT need
to be open; if no session is found this logs into Opus with the OO credential and
launches it, exactly as the robots do.

Usage:
    uv run python sandbox_layout.py
        END TO END. Claims a real window from dbo.CJI3_Udtraek, extracts it, loads it
        into dbo.CJI3 and reports. WRITES TO THE DATABASE - the window it processes is
        real work the backfill needs, and it closes it properly. On any failure after
        the claim, the window is handed back as Afventer.

    uv run python sandbox_layout.py --layout-only [01.06.2026 07.06.2026]
        Extract and report only. Touches no database, claims no window.

    uv run python sandbox_layout.py --file EXAMPLE.TXT
        Re-analyse an export that already exists. No SAP, no database.

Everything lands in sandbox_output/: run_log.txt, layout_report.txt,
spool_anonymised.txt and a screen dump per SAP screen.
"""

import os
import re
import sys
import uuid
from collections import Counter
from datetime import date, datetime, timedelta

# ---------------------------------------------------------------------------
# Run log - every step is timestamped, printed, and flushed to a file even when
# the run fails, so a failure leaves the same evidence a success does.
# ---------------------------------------------------------------------------

_LOG: list[str] = []


def log(message: str = "") -> None:
    """Print and record one line of the run log."""
    line = f"[{datetime.now():%H:%M:%S}] {message}" if message else ""
    print(line)
    _LOG.append(line)


def write_log() -> str:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, "run_log.txt")
    with open(path, "w", encoding="utf-8") as out:
        out.write("\n".join(_LOG) + "\n")
    return path

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.realpath(__file__)), "sandbox_output")
VARIANT_NAME = "FULDT UDTRÆK"
DYN_DATE_FIELD = "%%DYN002"
PLIST = "ROBOT"

# SAP's PRI_PARAMS-PRTXT field length. The spool overview sizes its Titel column to the
# content, so a long title is displayed in full and stays findable - the dispatcher has
# run with 22-character titles. This is only here to catch a runaway prefix.
MAX_PRTXT_LENGTH = 68

# Stop at the dynamic selections screen and wait for Enter. Off: the run is unattended
# and everything lands in sandbox_output/ instead.
PAUSE_AT_DYN_SELECTIONS = False

# The label %%DYN002 must carry. Confirmed from a screen dump; if SAP disagrees the
# dates would go into a different field and the extract would silently cover the wrong
# rows, so the run stops instead.
EXPECTED_DYN_LABEL = "Registreringsdato"

# Label beside R_BUDAT on the main selection screen, checked before it is
# overwritten with the window-relative posting-date range.
EXPECTED_BUDAT_LABEL = "Bogføringsdato"

# Env vars the OpenOrchestrator connection is built from, same as sandbox.py.
ENV_SQL = "OpenOrchestratorSQL"
ENV_KEY = "OpenOrchestratorKey"


# ---------------------------------------------------------------------------
# OpenOrchestrator + SAP startup
# ---------------------------------------------------------------------------

def get_orchestrator_connection():
    """Build an OO connection from the environment, as the robots get one from args."""
    from OpenOrchestrator.orchestrator_connection.connection import (  # noqa: E402  pylint: disable=import-outside-toplevel
        OrchestratorConnection,
    )

    conn_string, crypto_key = os.getenv(ENV_SQL), os.getenv(ENV_KEY)
    if not conn_string or not crypto_key:
        raise SystemExit(
            f"Set the {ENV_SQL} and {ENV_KEY} environment variables first - the SAP "
            "login credential is read from OpenOrchestrator, not stored here."
        )
    return OrchestratorConnection("SAPCJI3Layout", conn_string, crypto_key, None, None, None)


def find_sap_session():
    """Return the first SAP session if SAP is already running, else None."""
    import win32com.client  # pylint: disable=import-outside-toplevel

    try:
        application = win32com.client.GetObject("SAPGUI").GetScriptingEngine
        if application.Children.Count == 0:
            return None
        connection = application.Children(0)
        if connection.Children.Count == 0:
            return None
        return connection.Children(0)
    except Exception:  # pylint: disable=broad-exception-caught
        return None


def ensure_sap(orchestrator_connection):
    """Attach to a running SAP, or log into Opus and launch it the way the robots do."""
    session = find_sap_session()
    if session is not None:
        print("SAP is already running - attaching to the existing session.")
    else:
        print("No SAP session found. Logging into Opus and launching SAP...")
        from initialize_sap import initialize_sap  # pylint: disable=import-outside-toplevel

        if not initialize_sap(orchestrator_connection):
            raise SystemExit("SAP failed to launch.")
        session = find_sap_session()
        if session is None:
            raise SystemExit("SAP launched but no scripting session appeared.")
        print("SAP launched.")

    # Normalise to SAP Easy Access so the recorded CJI3 sequence starts where it expects.
    from initialize_sap import dismiss_until_easy_access  # pylint: disable=import-outside-toplevel

    try:
        dismiss_until_easy_access(30)
        print("SAP is at Easy Access.")
    except TimeoutError as error:
        print(f"WARNING: could not reach Easy Access ({error}).")
        print("         Navigate SAP to the main screen manually, then continue.")

    return session


# ---------------------------------------------------------------------------
# SAP screen diagnostics
# ---------------------------------------------------------------------------

def dump_screen(session, tag: str, mask_text: bool = False) -> str:
    """
    Write every control on the current screen to a file, and return the path.

    This is the raw material for working out screen structure: control id, type, text
    and screen position. Call it on SELECTION screens freely - they hold field labels
    and whatever we typed. Pass mask_text=True on any screen showing result data, which
    masks control text while leaving ids and positions intact.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, f"sap_screen_{tag}.txt")

    def prop(control, name, default=""):
        try:
            value = getattr(control, name)
            return default if value is None else value
        except Exception:  # pylint: disable=broad-exception-caught
            return default

    with open(path, "w", encoding="utf-8") as out:
        out.write(f"SAP SCREEN DUMP: {tag}\n")
        out.write("=" * 78 + "\n")
        try:
            window = session.findById("wnd[0]")
            out.write(f"window title : {window.Text!r}\n")
        except Exception as error:  # pylint: disable=broad-exception-caught
            out.write(f"window title : <unreadable: {error}>\n")
        try:
            out.write(f"status bar   : {session.findById('wnd[0]/sbar/pane[0]').Text!r}\n")
        except Exception:  # pylint: disable=broad-exception-caught
            pass
        out.write(f"text masked  : {mask_text}\n\n")

        try:
            usr = session.findById("wnd[0]/usr")
        except Exception as error:  # pylint: disable=broad-exception-caught
            out.write(f"no wnd[0]/usr on this screen: {error}\n")
            return path

        out.write(f"{'type':<16} {'left':>5} {'top':>4} {'w':>4}  {'text':<34} id\n")
        out.write("-" * 78 + "\n")
        rows = []
        for index in range(usr.Children.Count):
            control = usr.Children(index)
            text = str(prop(control, "Text"))
            rows.append((
                int(prop(control, "Top", 0)),
                int(prop(control, "Left", 0)),
                str(prop(control, "Type")),
                int(prop(control, "Width", 0)),
                mask(text) if mask_text else text,
                str(prop(control, "Id")),
            ))
        for top, left, ctype, width, text, cid in sorted(rows):
            short_id = cid.split("/usr/", 1)[-1]
            out.write(f"{ctype:<16} {left:>5} {top:>4} {width:>4}  {text[:34]:<34} {short_id}\n")

    print(f"  screen dump -> {path}")
    return path


def dump_menu(session, tag: str, max_depth: int = 4) -> str:
    """
    Walk the whole menu bar and write every entry with its id and text.

    Kept because it is how the export format was settled: menu[3] ('Tekst med
    tabulator') was found to be unparseable and menu[1] ('Eksporter som tekst')
    replaced it. Run it again if the submenu ever changes - guessing at menu ids from
    memory is how the ALV suggestion went wrong.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, f"sap_menu_{tag}.txt")

    def walk(node, node_id: str, depth: int, out) -> None:
        if depth > max_depth:
            return
        try:
            count = node.Children.Count
        except Exception:  # pylint: disable=broad-exception-caught
            return
        for index in range(count):
            try:
                child = node.Children(index)
                text = (child.Text or "").strip()
                child_id = f"{node_id}/menu[{index}]"
                out.write(f"{'  ' * depth}{child_id:<46} {text!r}\n")
                walk(child, child_id, depth + 1, out)
            except Exception as error:  # pylint: disable=broad-exception-caught
                out.write(f"{'  ' * depth}<unreadable at index {index}: {error}>\n")

    with open(path, "w", encoding="utf-8") as out:
        out.write(f"SAP MENU DUMP: {tag}\n")
        out.write("=" * 78 + "\n")
        out.write(f"window title : {_window_title(session)!r}\n\n")
        out.write("The export in use is menu[0]/menu[2]/menu[1] - Eksporter som tekst.\n")
        out.write("menu[3] (Tekst med tabulator) was rejected: not parseable.\n\n")
        try:
            mbar = session.findById("wnd[0]/mbar")
        except Exception as error:  # pylint: disable=broad-exception-caught
            out.write(f"no menu bar on this screen: {error}\n")
            return path
        walk(mbar, "mbar", 0, out)

    print(f"  menu dump -> {path}")
    return path


def read_dyn_field_label(session, field_id: str) -> str:
    """
    Label belonging to a dynamic selection field. Mirrors the dispatcher's copy.

    SAP gives the label its own control with an id derived from the field's, so
    %%DYN002 is described by txt%_%%DYN002_%_APP_%-TEXT. That is read first. The
    earlier version of this function only searched for GuiLabel controls by screen
    position and returned '' - on this screen the labels are GuiTextField, which is
    why it found nothing.
    """
    try:
        text = session.findById(f"wnd[0]/usr/txt%_{field_id}_%_APP_%-TEXT").Text
        if text and text.strip():
            return text.strip()
    except Exception:  # pylint: disable=broad-exception-caught
        pass  # Fall through to the positional search.

    field = session.findById(f"wnd[0]/usr/ctxt{field_id}-LOW")
    usr = session.findById("wnd[0]/usr")

    best_text, best_left = "", -1
    for index in range(usr.Children.Count):
        control = usr.Children(index)
        try:
            if control.Type not in ("GuiLabel", "GuiTextField"):
                continue
            if control.Top != field.Top or control.Left >= field.Left:
                continue
            if control.Left > best_left:
                best_left, best_text = control.Left, control.Text.strip()
        except Exception:  # pylint: disable=broad-exception-caught
            continue
    return best_text


# ---------------------------------------------------------------------------
# SAP: submit the extract
# ---------------------------------------------------------------------------

# Mirrors config.BUDAT_MONTHS_BACK in the dispatcher. 15 rather than the 14 the
# recording implies: Oekonomi confirmed a fiscal year stays open for posting until
# roughly 20 February of the following year, so a posting date can sit about 14 months
# behind its registration date. 15 leaves a month of margin.
BUDAT_MONTHS_BACK = 15


def budat_range(dato_fra, dato_til) -> tuple[str, str]:
    """
    Posting-date range for an entry-date window, as dd.mm.yyyy. Mirrors the dispatcher.

    low  = first day of the month, BUDAT_MONTHS_BACK months before dato_fra
    high = last day of the month containing dato_til

    Reproduces the manual procedure recorded on 16.02.2026 (01.12.2024 - 28.02.2026)
    and, for the week of 12-18 August 2026, the exact range the variant already stored.
    """
    months = dato_fra.year * 12 + (dato_fra.month - 1) - BUDAT_MONTHS_BACK
    low = date(months // 12, months % 12 + 1, 1)
    first_of_next = date(dato_til.year + (dato_til.month == 12),
                         (dato_til.month % 12) + 1, 1)
    return low.strftime("%d.%m.%Y"), (first_of_next - timedelta(days=1)).strftime("%d.%m.%Y")


def submit_cji3_extract(session, date_low: str, date_high: str, prtxt: str,
                        budat_dates: tuple[str, str] | None = None) -> dict:
    """
    Run CJI3 for a date range and print it to a spool job named prtxt.

    Returns {'label': ..., 'status_after_execute': ..., 'title_after_execute': ...}.

    That status bar reading is the point of interest: when the selection matches no
    postings, the background job still completes but produces NO spool job, so the
    performer would poll for something that will never appear. SAP announces the empty
    result here, right after Execute, which is the only place it can be caught cheaply.
    """
    from robot_framework.process import wait_ready  # pylint: disable=import-outside-toplevel

    session.findById("wnd[0]/tbar[0]/okcd").text = "CJI3"
    session.findById("wnd[0]").sendVKey(0)
    wait_ready(session)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()          # Hent variant
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/usr/txtV-LOW").text = VARIANT_NAME
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    wait_ready(session)

    dump_screen(session, "01_selection_screen")

    # Overwrite the variant's stored Bogfoeringsdato range, relative to this window.
    # Mirrors budat_range in the dispatcher; see its config for the rule and why the
    # variant's own values cannot be trusted.
    if budat_dates is not None:
        budat_low, budat_high = budat_dates
        budat_label = read_dyn_field_label(session, "R_BUDAT")
        log(f"  R_BUDAT label={budat_label!r}, was "
            f"{session.findById('wnd[0]/usr/ctxtR_BUDAT-LOW').Text!r} - "
            f"{session.findById('wnd[0]/usr/ctxtR_BUDAT-HIGH').Text!r}")
        if budat_label != EXPECTED_BUDAT_LABEL:
            raise SystemExit(
                f"Expected R_BUDAT to be labelled {EXPECTED_BUDAT_LABEL!r}, got "
                f"{budat_label!r}."
            )
        session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = budat_low
        session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = budat_high
        log(f"  R_BUDAT set to {budat_low} - {budat_high}")

    session.findById("wnd[0]").sendVKey(21)                     # Dynamiske selektioner
    wait_ready(session)

    dump_screen(session, "02_dynamic_selections")

    # =======================================================================
    # >>> PUT YOUR BREAKPOINT ON THE NEXT LINE <<<
    #
    # This is the moment the dynamic selections screen is open, before any date
    # is typed. %%DYN002 is a POSITION, not a field name - SAP numbers these
    # fields by the order the variant puts them in - so this is where we find out
    # which CJI3 date the whole extract is actually filtered on.
    #
    # When it stops, try these in the debug console:
    #
    #   read_dyn_field_label(session, "%%DYN002")
    #       -> the label sitting left of the field, e.g. 'Bogføringsdato'
    #
    #   session.findById("wnd[0]/usr/ctxt%%DYN002-LOW").Text
    #   session.findById("wnd[0]/usr/ctxt%%DYN002-LOW").Tooltip
    #       -> the tooltip often carries the technical name (BUDAT/CPUDT/BLDAT)
    #
    #   dump_screen(session, "manual", mask_text=False)
    #       -> writes every control with position and text to sandbox_output/
    #
    # Also worth checking: is %%DYN001 or %%DYN003 the date you actually want?
    #   read_dyn_field_label(session, "%%DYN001")
    #   read_dyn_field_label(session, "%%DYN003")
    #
    # sandbox_output/02_dynamic_selections.txt already holds the same dump, so if
    # you would rather not use the debugger at all, just read that file.
    # =======================================================================
    label = read_dyn_field_label(session, DYN_DATE_FIELD)
    print(f"  {DYN_DATE_FIELD} label -> {label!r}")

    if PAUSE_AT_DYN_SELECTIONS:
        input("  Dynamic selections open. Inspect SAP, then press Enter to continue...")

    session.findById(f"wnd[0]/usr/ctxt{DYN_DATE_FIELD}-LOW").text = date_low
    session.findById(f"wnd[0]/usr/ctxt{DYN_DATE_FIELD}-HIGH").text = date_high

    session.findById("wnd[0]/tbar[0]/btn[11]").press()          # Udfoer
    wait_ready(session)

    # Everything on screen now is real result data, so this dump is masked.
    dump_screen(session, "03_after_execute", mask_text=True)
    status_after_execute = _status_bar(session)
    title_after_execute = _window_title(session)
    log(f"  after Execute: title={title_after_execute!r}")
    log(f"  after Execute: status bar={status_after_execute!r}")

    session.findById("wnd[0]/mbar/menu[0]/menu[2]").select()    # Udskriv
    session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").text = PLIST
    session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").text = prtxt
    session.findById("wnd[1]/tbar[0]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSOFORT_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]/tbar[0]/btn[15]").press()          # Afslut
    wait_ready(session)

    return {
        "label": label,
        "status_after_execute": status_after_execute,
        "title_after_execute": title_after_execute,
    }


def unique_prtxt(prefix: str) -> str:
    """
    Build a spool job title that is unique to this run.

    _find_spool_job_row matches the title as a SUBSTRING of what the spool overview
    displays, so a title that repeats matches an EARLIER spool job just as well as the
    new one - and the robot then exports a previous run's data and reports on stale rows.
    'LAYOUT{today}' did exactly that: two identical LAYOUT20260819 jobs were sitting in
    the overview, and a second run that day would have found the first one's output.

    The SANDBOX prefix also keeps these distinguishable from the dispatcher's real jobs
    in SP01, which use YYYYMMDDUGEww_<random>.
    """
    prtxt = f"{prefix}{uuid.uuid4().hex[:6]}"
    if len(prtxt) > MAX_PRTXT_LENGTH:
        raise ValueError(
            f"Spool job title {prtxt!r} is {len(prtxt)} characters, over the "
            f"{MAX_PRTXT_LENGTH} the spool overview is known to render. Shorten the "
            "prefix."
        )
    return prtxt


def _status_bar(session) -> str:
    try:
        return (session.findById("wnd[0]/sbar/pane[0]").Text or "").strip()
    except Exception:  # pylint: disable=broad-exception-caught
        return ""


def _window_title(session) -> str:
    try:
        return (session.findById("wnd[0]").Text or "").strip()
    except Exception:  # pylint: disable=broad-exception-caught
        return ""


# ---------------------------------------------------------------------------
# Anonymisation and reporting
# ---------------------------------------------------------------------------

def mask(value: str) -> str:
    """Replace letters with a/A and digits with 9, keeping everything else."""
    out = []
    for char in value:
        if char.isdigit():
            out.append("9")
        elif char.isalpha():
            out.append("A" if char.isupper() else "a")
            
        else:
            out.append(char)
    return "".join(out)


def classify(values: list[str]) -> str:
    """Best-effort label for what a column holds, from the real values."""
    filled = [v for v in values if v]
    if not filled:
        return "always empty"
    if all(re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", v) for v in filled):
        return "date dd.mm.yyyy"
    if all(re.fullmatch(r"\d{2}:\d{2}:\d{2}", v) for v in filled):
        return "time hh:mm:ss"
    if all(re.fullmatch(r"-?[\d.]*\d(,\d+)?-?", v) for v in filled):
        return "danish number"
    shapes = {"date" if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", v)
              else "number" if re.fullmatch(r"-?[\d.]*\d(,\d+)?-?", v)
              else "text" for v in filled}
    return "MIXED: " + "/".join(sorted(shapes)) if len(shapes) > 1 else "text"


def analyse(raw_path: str, label: str, date_low: str, date_high: str, prtxt: str) -> None:
    """
    Write the anonymised copy and the layout report for an exported spool file.

    Columns come from the same offset slicing the real parser uses - imported, not
    reimplemented, so the report cannot disagree with what actually gets loaded.

    The anonymised copy needs no splitting at all: mask() only touches letters and
    digits, so pipes, padding and line lengths survive untouched and the copy is a
    faithful stand-in for the original.
    """
    from robot_framework.spool_to_sql import (  # pylint: disable=import-outside-toplevel
        _column_ranges, _is_data_row, _is_header_line, _is_separator_line, _slice_all,
    )

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(raw_path, encoding="cp1252", errors="replace") as handle:
        lines = [line.rstrip("\r\n") for line in handle]

    header_idx = [i for i, line in enumerate(lines) if _is_header_line(line)]
    if not header_idx:
        raise SystemExit("No header line containing 'Bilagsnummer' - cannot analyse.")

    first = header_idx[0]
    ranges = _column_ranges(lines[first])
    names = _slice_all(lines[first], ranges)
    sub_line = lines[first + 1] if first + 1 < len(lines) else ""
    sub_ranges = _column_ranges(sub_line) if "|" in sub_line else []
    sub_names = _slice_all(sub_line, sub_ranges) if sub_ranges else []

    bilag_idx = names.index("Bilagsnummer")
    data_rows = [l for l in lines
                 if not _is_header_line(l) and not _is_separator_line(l)
                 and _is_data_row(_slice_all(l, ranges), bilag_idx)]

    # ---- anonymised copy: header lines verbatim, every other line masked in place ----
    anon_path = os.path.join(OUTPUT_DIR, "spool_anonymised.txt")
    with open(anon_path, "w", encoding="cp1252", errors="replace") as out:
        for i, line in enumerate(lines):
            keep = i in header_idx or (i - 1) in header_idx
            out.write((line if keep else mask(line)) + "\n")

    # ---- layout report ----
    report_path = os.path.join(OUTPUT_DIR, "layout_report.txt")
    lengths = Counter(len(l) for l in data_rows)
    header_bars = [i for i, c in enumerate(lines[first]) if c == "|"]
    off_layout = sum(1 for l in data_rows
                     if [i for i, c in enumerate(l) if c == "|"] != header_bars)

    with open(report_path, "w", encoding="utf-8") as out:
        out.write("CJI3 LAYOUT REPORT\n")
        out.write("=" * 78 + "\n")
        out.write(f"range                : {date_low} - {date_high}\n")
        out.write(f"spool job            : {prtxt}\n")
        out.write(f"{DYN_DATE_FIELD} label        : {label!r}\n")
        out.write(f"source file          : {os.path.basename(raw_path)}\n")
        out.write(f"total lines          : {len(lines)}\n")
        out.write(f"header lines (pages) : {len(header_idx)}\n")
        out.write(f"main columns         : {len(ranges)}\n")
        out.write(f"continuation columns : {len(sub_ranges)}\n")
        out.write(f"data rows            : {len(data_rows)}\n")
        out.write(f"main line lengths    : {dict(sorted(lengths.items()))}\n")
        out.write(f"rows with a pipe inside a value: {off_layout}\n")
        out.write("  (harmless - columns are sliced by offset, not split on the pipe)\n")

        out.write("\nMAIN COLUMNS, VERBATIM\n")
        out.write("-" * 78 + "\n")
        for i, (name, (a, b)) in enumerate(zip(names, ranges)):
            out.write(f"  [{i:>2}] chars {a:>4}-{str(b):<5} {name!r}\n")
        if sub_names:
            out.write("\nCONTINUATION COLUMNS, VERBATIM\n")
            out.write("-" * 78 + "\n")
            for i, (name, (a, b)) in enumerate(zip(sub_names, sub_ranges)):
                out.write(f"  [{i:>2}] chars {a:>4}-{str(b):<5} {name!r}\n")

        out.write("\nPER-COLUMN MEASUREMENTS (from real values, none reproduced)\n")
        out.write("-" * 78 + "\n")
        out.write(f"{'idx':>4} {'heading':<28} {'max':>4} {'distinct':>8}  "
                  f"{'kind':<16} longest (masked)\n")
        sliced = [_slice_all(l, ranges) for l in data_rows]
        for i, name in enumerate(names):
            if not name:
                continue
            values = [row[i] for row in sliced]
            longest = max(values, key=len) if values else ""
            out.write(f"{i:>4} {name[:28]:<28} {len(longest):>4} "
                      f"{len(set(v for v in values if v)):>8}  "
                      f"{classify(values):<16} {mask(longest)[:24]}\n")

    print(f"\nWrote:\n  {anon_path}\n  {report_path}")
    print("\nSafe to share: headings are metadata, every value is masked.")


# ---------------------------------------------------------------------------
# Work list access. Mirrors udtraek.py in the dispatcher repo; duplicated because the
# two repos cannot import from each other.
# ---------------------------------------------------------------------------

def _result_rows(cursor) -> list[dict]:
    names = [column[0] for column in cursor.description]
    return [dict(zip(names, row)) for row in cursor.fetchall()]


def verify_schema(conn) -> bool:
    """Confirm the DDL actually created everything, before anything depends on it."""
    expected_tables = ["CJI3", "CJI3_Loadlog", "CJI3_Stage", "CJI3_Udtraek"]
    expected_procs = ["usp_CJI3_AfslutUdtraek", "usp_CJI3_Merge",
                      "usp_CJI3_OpdaterVindueliste", "usp_CJI3_ReserverUdtraek",
                      "usp_CJI3_SeedUdtraek", "usp_CJI3_TilknytSpoolJob"]
    expected_views = ["vw_CJI3_DageUdenPoster", "vw_CJI3_Fremdrift", "vw_CJI3_Udtraekstatus"]

    cursor = conn.cursor()
    cursor.execute("""
        SELECT name, type_desc FROM sys.objects
        WHERE type_desc IN ('USER_TABLE', 'SQL_STORED_PROCEDURE', 'VIEW',
                            'SQL_SCALAR_FUNCTION')
          AND (name LIKE '%CJI3%' OR name = 'fn_DkDecimal')
    """)
    found = {name for name, _ in cursor.fetchall()}

    ok = True
    for label, expected in (("table", expected_tables), ("procedure", expected_procs),
                            ("view", expected_views), ("function", ["fn_DkDecimal"])):
        missing = [n for n in expected if n not in found]
        if missing:
            ok = False
            log(f"  MISSING {label}(s): {missing}")
        else:
            log(f"  all {len(expected)} {label}(s) present")

    # Compare column NAMES, not just the count. A count alone is not enough: removing
    # Kildefil and adding DataTidspunkt left the total unchanged at 69, so a stale table
    # with the wrong set of columns would have passed. The expected names are read from
    # create_CJI3.sql itself, so this check maintains itself.
    cursor.execute("""
        SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'CJI3'
    """)
    live = {row[0] for row in cursor.fetchall()}
    log(f"  dbo.CJI3 has {len(live)} columns")

    declared = _declared_cji3_columns()
    if declared:
        missing = sorted(declared - live)
        extra = sorted(live - declared)
        if missing or extra:
            ok = False
            if missing:
                log(f"  MISSING from the table: {missing}")
            if extra:
                log(f"  NOT IN create_CJI3.sql: {extra}")
            log("  The tables in create_CJI3.sql are guarded by IF OBJECT_ID IS NULL, so "
                "re-running that script does NOT change an existing table. Run "
                "sql/drop_CJI3.sql first, then create_CJI3.sql.")
        else:
            log(f"  all {len(declared)} declared columns present, none extra")
    return ok


def _declared_cji3_columns() -> set[str]:
    """Column names create_CJI3.sql declares for dbo.CJI3, or an empty set if unreadable."""
    path = os.path.join(os.path.dirname(os.path.realpath(__file__)),
                        "sql", "create_CJI3.sql")
    try:
        sql = open(path, encoding="utf-8").read()
        start = sql.index("CREATE TABLE dbo.CJI3" + chr(10))
        block = sql[start:sql.index("CONSTRAINT PK_CJI3 ", start)]
    except (OSError, ValueError) as error:
        log(f"  (could not read create_CJI3.sql to compare columns: {error})")
        return set()
    return set(re.findall(r"^\s{8}([A-Za-z][A-Za-z0-9]*)\s+[A-Z]", block, re.M))


def opdater_vindueliste(conn) -> dict:
    cursor = conn.cursor()
    cursor.execute("{CALL dbo.usp_CJI3_OpdaterVindueliste (?, ?)}", 7, 7)
    rows = _result_rows(cursor)
    return rows[0] if rows else {}


def reserver_udtraek(conn, antal: int) -> list[dict]:
    cursor = conn.cursor()
    cursor.execute("{CALL dbo.usp_CJI3_ReserverUdtraek (?)}", antal)
    return _result_rows(cursor)


def tilknyt_spooljob(conn, udtraek_id: int, spool_job: str) -> None:
    conn.cursor().execute("{CALL dbo.usp_CJI3_TilknytSpoolJob (?, ?)}",
                          udtraek_id, spool_job)


def afslut_udtraek(conn, udtraek_id: int, status: str, message: str | None) -> None:
    conn.cursor().execute("{CALL dbo.usp_CJI3_AfslutUdtraek (?, ?, ?)}",
                          udtraek_id, status, (message or "")[:1000] or None)


# ---------------------------------------------------------------------------
# End-to-end run
# ---------------------------------------------------------------------------

def run_end_to_end() -> None:
    """
    Claim a real window, extract it, load it, and report - the whole pipeline minus
    the queue hop through the queuer and performer.

    This WRITES to BI_Oekonomi: it claims a window, inserts rows into dbo.CJI3 and
    closes the window. That is real work the backfill needs done, not throwaway. If
    anything fails after the claim, the window is handed back as Afventer so the next
    run retries it rather than leaving it stuck IGang.
    """
    from robot_framework import mssql_load  # pylint: disable=import-outside-toplevel
    from robot_framework.process import (  # pylint: disable=import-outside-toplevel
        export_spool_as_text, get_exported_file_path, open_spool_overview,
        select_spool_job, wait_for_session, wait_for_spool_job, wait_ready,
    )

    log("=" * 74)
    log("CJI3 END-TO-END SANDBOX")
    log("=" * 74)

    orchestrator_connection = get_orchestrator_connection()
    log("OpenOrchestrator connection open.")

    conn = mssql_load.connect(orchestrator_connection)
    conn.autocommit = True
    log(f"Connected to {mssql_load.config.SQL_DATABASE} "
        f"via {mssql_load.config.SQL_DRIVER}.")

    window = None
    try:
        log("")
        log("STEP 1 - schema")
        if not verify_schema(conn):
            raise SystemExit("Schema is not as expected; see the note above.")

        log("")
        log("STEP 2 - work list")
        summary = opdater_vindueliste(conn)
        log(f"  OpdaterVindueliste -> {summary}")
        windows = reserver_udtraek(conn, 1)
        if not windows:
            log("  nothing pending; nothing to do.")
            return
        window = windows[0]
        log(f"  claimed window {window['UdtraekId']}: "
            f"{window['DatoFraSAP']} - {window['DatoTilSAP']} "
            f"(forsoeg {window['Forsoeg']}, koersler {window['AntalKoersler']})")

        log("")
        log("STEP 3 - SAP")
        session = ensure_sap(orchestrator_connection)
        wait_ready(session)
        session.findById("wnd[0]").maximize()

        prtxt = unique_prtxt(f"SANDBOX_W{window['UdtraekId']}_")
        budat = budat_range(window["DatoFra"], window["DatoTil"])
        result = submit_cji3_extract(session, window["DatoFraSAP"],
                                    window["DatoTilSAP"], prtxt, budat_dates=budat)
        label = result["label"]
        log(f"  {DYN_DATE_FIELD} label read as {label!r}")
        if label != EXPECTED_DYN_LABEL:
            raise SystemExit(
                f"Expected {DYN_DATE_FIELD} to be {EXPECTED_DYN_LABEL!r} but SAP says "
                f"{label!r}. The variant's dynamic selections have changed - the dates "
                "would filter the wrong field."
            )
        log(f"  label matches {EXPECTED_DYN_LABEL!r} - guard passes")
        log(f"  submitted to spool as {prtxt}")

        tilknyt_spooljob(conn, window["UdtraekId"], prtxt)
        log("  spool job recorded on the window")

        log("")
        log("STEP 4 - spool")
        open_spool_overview(session)
        spool_session = wait_for_session(connection_index=0, session_index=1)
        wait_ready(spool_session)
        dump_screen(spool_session, "05_spool_overview")
        dump_menu(spool_session, "spool_overview")
        started = datetime.now()
        row = wait_for_spool_job(spool_session, prtxt)
        log(f"  ready at row {row} after {(datetime.now() - started).seconds}s")
        select_spool_job(spool_session, row)
        export_spool_as_text(spool_session)
        raw_path = get_exported_file_path(spool_session)
        log(f"  exported to {raw_path}")

        log("")
        log("STEP 5 - layout report")
        analyse(raw_path, label, window["DatoFraSAP"], window["DatoTilSAP"], prtxt)

        log("")
        log("STEP 6 - load into CJI3")
        counts = mssql_load.load_spool_file(
            orchestrator_connection, file_path=raw_path,
            spool_job=prtxt, udtraek_id=window["UdtraekId"])
        for key, value in counts.items():
            log(f"  {key:<24} {value}")

        log("")
        log("STEP 7 - read back")
        cursor = conn.cursor()
        cursor.execute("SELECT TOP 1 * FROM dbo.vw_CJI3_Udtraekstatus "
                       "WHERE UdtraekId = ?", window["UdtraekId"])
        for key, value in (_result_rows(cursor) or [{}])[0].items():
            log(f"  {key:<24} {value}")
        cursor.execute("SELECT COUNT(*) FROM dbo.CJI3")
        log(f"  rows now in dbo.CJI3     {cursor.fetchone()[0]}")
        cursor.execute("SELECT COUNT(*) FROM dbo.CJI3_Stage")
        log(f"  rows left in staging     {cursor.fetchone()[0]} (should be 0)")

        log("")
        log("SUCCESS - the full path works end to end.")

    except BaseException as error:  # noqa: BLE001  pylint: disable=broad-exception-caught
        log("")
        log(f"FAILED: {type(error).__name__}: {error}")
        if window is not None:
            try:
                afslut_udtraek(conn, window["UdtraekId"], "Afventer",
                               f"sandbox_layout: {type(error).__name__}: {error}")
                log(f"  window {window['UdtraekId']} handed back as Afventer")
            except Exception as release_error:  # pylint: disable=broad-exception-caught
                log(f"  could not hand the window back: {release_error}")
        raise
    finally:
        conn.close()
        log("")
        log(f"run log -> {write_log()}")


def try_text_export(spool_job: str) -> None:
    """
    Select an export format on an EXISTING spool job and dump whatever appears.

    This is how the format was settled: menu[3] ('Tekst med tabulator') produced a file
    that could not be parsed reliably, and menu[1] ('Eksporter som tekst') - pipe
    delimited fixed width - parsed with zero bad values on the same data. Kept for the
    next time a menu path needs checking rather than guessed at.

    Nothing is consumed: exporting does not delete the spool job.
    """
    from robot_framework.process import (  # pylint: disable=import-outside-toplevel
        open_spool_overview, select_spool_job, wait_for_session, wait_for_spool_job,
        wait_ready,
    )

    log("=" * 74)
    log(f"TRY EXPORT FORMAT ON SPOOL JOB {spool_job}")
    log("=" * 74)

    orchestrator_connection = get_orchestrator_connection()
    session = ensure_sap(orchestrator_connection)
    wait_ready(session)
    session.findById("wnd[0]").maximize()

    open_spool_overview(session)
    spool_session = wait_for_session(connection_index=0, session_index=1)
    wait_ready(spool_session)

    dump_menu(spool_session, "spool_overview")

    row = wait_for_spool_job(spool_session, spool_job)
    log(f"found {spool_job} at row {row}")
    select_spool_job(spool_session, row)

    log("selecting menu[0]/menu[2]/menu[1] - Eksporter som tekst")
    spool_session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[1]").select()
    wait_ready(spool_session)

    log(f"after select: wnd[0] title={_window_title(spool_session)!r}")
    log(f"after select: status bar={_status_bar(spool_session)!r}")
    dump_screen(spool_session, "06_export_dialog")

    for window_index in (1, 2):
        try:
            popup = spool_session.findById(f"wnd[{window_index}]")
        except Exception:  # pylint: disable=broad-exception-caught
            continue
        log(f"popup wnd[{window_index}] present: {popup.Text!r}")

    log("")
    log(f"run log -> {write_log()}")


def main() -> None:
    if len(sys.argv) >= 3 and sys.argv[1] == "--try-export":
        try_text_export(sys.argv[2])
        return

    if len(sys.argv) >= 3 and sys.argv[1] == "--file":
        analyse(sys.argv[2], label="(not read - SAP was not opened)",
                date_low="?", date_high="?", prtxt="(existing file)")
        return

    if len(sys.argv) >= 2 and sys.argv[1] == "--layout-only":
        run_layout_only()
        return

    run_end_to_end()


def run_layout_only() -> None:
    """Original behaviour: extract a date range and report on it, without touching SQL."""

    from robot_framework.process import (  # pylint: disable=import-outside-toplevel
        export_spool_as_text,
        get_exported_file_path,
        open_spool_overview,
        select_spool_job,
        wait_for_session,
        wait_for_spool_job,
        wait_ready,
    )

    # argv[1] is --layout-only, so any dates follow it.
    if len(sys.argv) == 4:
        date_low, date_high = sys.argv[2], sys.argv[3]
    else:
        high = date.today() - timedelta(days=1)
        low = high - timedelta(days=6)
        date_low, date_high = low.strftime("%d.%m.%Y"), high.strftime("%d.%m.%Y")

    prtxt = unique_prtxt("SANDBOX_LAY_")
    print(f"Range {date_low} - {date_high}, spool job {prtxt}\n")

    orchestrator_connection = get_orchestrator_connection()
    session = ensure_sap(orchestrator_connection)
    wait_ready(session)
    session.findById("wnd[0]").maximize()

    label = submit_cji3_extract(session, date_low, date_high, prtxt)["label"]

    print("\nWaiting for the spool job (the slow part - up to 30 minutes)...")
    open_spool_overview(session)
    spool_session = wait_for_session(connection_index=0, session_index=1)
    wait_ready(spool_session)

    dump_screen(spool_session, "03_spool_overview")

    row = wait_for_spool_job(spool_session, prtxt)
    print(f"  spool job ready at row {row}")
    select_spool_job(spool_session, row)
    export_spool_as_text(spool_session)
    raw_path = get_exported_file_path(spool_session)
    print(f"  exported to {raw_path}")

    # >>> SECOND USEFUL BREAKPOINT <<<
    # Stop here to look at the raw export before it is masked. raw_path holds the
    # real file; sandbox_output/ will hold only the anonymised copy.
    analyse(raw_path, label, date_low, date_high, prtxt)


if __name__ == "__main__":
    main()
