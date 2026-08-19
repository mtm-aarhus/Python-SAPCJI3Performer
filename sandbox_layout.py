"""Debuggable diagnostic: launch SAP if needed, dispatch a CJI3 extract, wait for the
spool job, and write an ANONYMISED copy of the raw export plus a layout report.

Press F5 in VS Code ("Debug: CJI3 layout sandbox"). Nothing here touches SQL and no
queue elements are created - it only drives SAP and writes files.

What it answers, which the old EXAMPLE.TXT cannot:
  * the exact raw column headings of the CURRENT layout, verbatim
  * how wide values in each column actually get, so create_CJI3.sql can be sized
    from fact instead of inference
  * which CJI3 date field %%DYN002 really is - it reads the on-screen label
  * whether the tab layout is positionally reliable across rows

Anonymisation: every letter becomes 'a'/'A' and every digit '9', in place. Lengths,
punctuation, tabs and case pattern survive; content does not. There is no key and no
cipher, so nothing can be recovered. A date still reads '99.99.9999' and an amount
'9.999,99', which is all that is needed to judge layout. Column HEADINGS are kept
verbatim - they are metadata, not data.

Prerequisites: the two OpenOrchestrator environment variables below. SAP does NOT need
to be open; if no session is found this logs into Opus with the OO credential and
launches it, exactly as the robots do.

Usage:
    uv run python sandbox_layout.py                     # last full week
    uv run python sandbox_layout.py 01.06.2026 07.06.2026
    uv run python sandbox_layout.py --file EXAMPLE.TXT   # analyse only, no SAP
"""

import os
import re
import sys
from collections import Counter
from datetime import date, timedelta

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.realpath(__file__)), "sandbox_output")
VARIANT_NAME = "FULDT UDTRÆK"
DYN_DATE_FIELD = "%%DYN002"
PLIST = "ROBOT"

# Stop at the dynamic selections screen and wait for Enter, so the screen can be
# inspected by eye even when running without a debugger. Set False for an unattended run.
PAUSE_AT_DYN_SELECTIONS = True

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


def read_dyn_field_label(session, field_id: str) -> str:
    """Label immediately left of a dynamic selection field, matched by screen position."""
    field = session.findById(f"wnd[0]/usr/ctxt{field_id}-LOW")
    usr = session.findById("wnd[0]/usr")

    best_text, best_left = "", -1
    for index in range(usr.Children.Count):
        control = usr.Children(index)
        try:
            if control.Type != "GuiLabel":
                continue
            if control.Top != field.Top or control.Left >= field.Left:
                continue
            if control.Left > best_left:
                best_left, best_text = control.Left, control.Text.strip()
        except Exception:  # pylint: disable=broad-exception-caught
            continue

    if not best_text:
        try:
            best_text = (field.Tooltip or "").strip()
        except Exception:  # pylint: disable=broad-exception-caught
            best_text = ""
    return best_text


# ---------------------------------------------------------------------------
# SAP: submit the extract
# ---------------------------------------------------------------------------

def submit_cji3_extract(session, date_low: str, date_high: str, prtxt: str) -> str:
    """Run CJI3 for a date range, print to spool as prtxt, return the date field label."""
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

    session.findById("wnd[0]/mbar/menu[0]/menu[2]").select()    # Udskriv
    session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").text = PLIST
    session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").text = prtxt
    session.findById("wnd[1]/tbar[0]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSOFORT_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]/tbar[0]/btn[15]").press()          # Afslut
    wait_ready(session)

    return label


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
    """Write the anonymised copy and the layout report for an exported spool file."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(raw_path, encoding="cp1252", errors="replace") as handle:
        lines = [line.rstrip("\r\n") for line in handle]

    header_idx = [i for i, line in enumerate(lines)
                  if "Bilagsnummer" in [f.strip() for f in line.split("\t")]]
    if not header_idx:
        raise SystemExit("No header line containing 'Bilagsnummer' - cannot analyse.")

    header = lines[header_idx[0]].split("\t")
    sub_header = lines[header_idx[0] + 1].split("\t") if header_idx[0] + 1 < len(lines) else []

    bilag_idx = [f.strip() for f in header].index("Bilagsnummer")
    data_rows = []
    for line in lines:
        fields = line.split("\t")
        if bilag_idx < len(fields) and re.fullmatch(r"\d+", fields[bilag_idx].strip()):
            data_rows.append(fields)

    # ---- anonymised copy: headers verbatim, everything else masked ----
    # cp1252, matching what SAP writes: the copy has to be a faithful stand-in for a
    # real export, so parse_spool_file can be tested against it directly. Writing UTF-8
    # here would mangle every ae/oe/aa in the headings when read back.
    anon_path = os.path.join(OUTPUT_DIR, "spool_anonymised.txt")
    with open(anon_path, "w", encoding="cp1252", errors="replace") as out:
        for i, line in enumerate(lines):
            if i in header_idx or (i - 1) in header_idx:
                out.write(line + "\n")
            else:
                out.write("\t".join(mask(f) for f in line.split("\t")) + "\n")

    # ---- layout report ----
    counts = Counter(len(r) for r in data_rows)
    report_path = os.path.join(OUTPUT_DIR, "layout_report.txt")
    with open(report_path, "w", encoding="utf-8") as out:
        out.write("CJI3 LAYOUT REPORT\n")
        out.write("=" * 78 + "\n")
        out.write(f"range                : {date_low} - {date_high}\n")
        out.write(f"spool job            : {prtxt}\n")
        out.write(f"{DYN_DATE_FIELD} label{' ' * 8}: {label!r}\n")
        out.write(f"source file          : {os.path.basename(raw_path)}\n")
        out.write(f"total lines          : {len(lines)}\n")
        out.write(f"header lines at      : {header_idx}\n")
        out.write(f"header field count   : {len(header)}\n")
        out.write(f"sub-header fields    : {len(sub_header)}\n")
        out.write(f"data rows            : {len(data_rows)}\n")
        out.write(f"data field counts    : {dict(sorted(counts.items()))}\n")

        # The single most important line in this report. Fewer fields than the header is
        # harmless if consistent - SAP trims trailing empty columns, so only the tail is
        # absent and every earlier index still lines up. Rows DIFFERING FROM EACH OTHER
        # is the dangerous case: at least one row is offset, and every column after the
        # offset holds its neighbour's value.
        if len(counts) > 1:
            out.write("\n  *** WARNING: data rows do not all have the same field count.\n")
            out.write("  Parsing by header index is unsafe here - rows with different\n")
            out.write("  counts are offset against each other, so values land in the\n")
            out.write("  wrong columns. Check the per-column 'kind' below: a date in a\n")
            out.write("  description column, or a PSP element in a unit-of-measure\n")
            out.write("  column, confirms it.\n")
        elif counts and len(header) not in counts:
            out.write("\n  OK: every data row has the same field count. It is lower than the\n")
            out.write(f"  header's {len(header)} because SAP trims trailing empty columns;\n")
            out.write("  earlier indices still align, so only the last columns are absent.\n")

        out.write("\nRAW HEADER, VERBATIM (one per line, with its index)\n")
        out.write("-" * 78 + "\n")
        for i, name in enumerate(header):
            out.write(f"  [{i:>2}] {name!r}\n")
        if sub_header:
            out.write("\nRAW SUB-HEADER (continuation line), VERBATIM\n")
            out.write("-" * 78 + "\n")
            for i, name in enumerate(sub_header):
                out.write(f"  [{i:>2}] {name!r}\n")

        out.write("\nPER-COLUMN MEASUREMENTS (from real values, none reproduced)\n")
        out.write("-" * 78 + "\n")
        out.write(f"{'idx':>4} {'heading':<28} {'max':>4} {'distinct':>8}  "
                  f"{'kind':<16} longest (masked)\n")
        for i, name in enumerate(header):
            if not name.strip():
                continue
            values = [r[i].strip() if i < len(r) else "" for r in data_rows]
            longest = max(values, key=len) if values else ""
            out.write(f"{i:>4} {name.strip()[:28]:<28} {len(longest):>4} "
                      f"{len(set(v for v in values if v)):>8}  "
                      f"{classify(values):<16} {mask(longest)[:24]}\n")

    print(f"\nWrote:\n  {anon_path}\n  {report_path}")
    print("\nSafe to share: headings are metadata, every value is masked.")


def main() -> None:
    if len(sys.argv) >= 3 and sys.argv[1] == "--file":
        analyse(sys.argv[2], label="(not read - SAP was not opened)",
                date_low="?", date_high="?", prtxt="(existing file)")
        return

    from robot_framework.process import (  # pylint: disable=import-outside-toplevel
        export_spool_as_tabtext,
        get_exported_file_path,
        open_spool_overview,
        select_spool_job,
        wait_for_session,
        wait_for_spool_job,
        wait_ready,
    )

    if len(sys.argv) == 3:
        date_low, date_high = sys.argv[1], sys.argv[2]
    else:
        high = date.today() - timedelta(days=1)
        low = high - timedelta(days=6)
        date_low, date_high = low.strftime("%d.%m.%Y"), high.strftime("%d.%m.%Y")

    prtxt = f"LAYOUT{date.today():%Y%m%d}"
    print(f"Range {date_low} - {date_high}, spool job {prtxt}\n")

    orchestrator_connection = get_orchestrator_connection()
    session = ensure_sap(orchestrator_connection)
    wait_ready(session)
    session.findById("wnd[0]").maximize()

    label = submit_cji3_extract(session, date_low, date_high, prtxt)

    print("\nWaiting for the spool job (the slow part - up to 30 minutes)...")
    open_spool_overview(session)
    spool_session = wait_for_session(connection_index=0, session_index=1)
    wait_ready(spool_session)

    dump_screen(spool_session, "03_spool_overview")

    row = wait_for_spool_job(spool_session, prtxt)
    print(f"  spool job ready at row {row}")
    select_spool_job(spool_session, row)
    export_spool_as_tabtext(spool_session)
    raw_path = get_exported_file_path(spool_session)
    print(f"  exported to {raw_path}")

    # >>> SECOND USEFUL BREAKPOINT <<<
    # Stop here to look at the raw export before it is masked. raw_path holds the
    # real file; sandbox_output/ will hold only the anonymised copy.
    analyse(raw_path, label, date_low, date_high, prtxt)


if __name__ == "__main__":
    main()
