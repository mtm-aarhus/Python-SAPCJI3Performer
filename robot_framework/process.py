"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement

import json
import os
import re
import time

import pywintypes
import win32com.client

from robot_framework import config
from robot_framework import mssql_load
from robot_framework.exceptions import BusinessError


_SAP_LABEL_ID_PATTERN = re.compile(r"/lbl\[(\d+),(\d+)\]$")
_SAP_CHECKBOX_ID_PATTERN = re.compile(r"/chk\[(\d+),(\d+)\]$")


# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    orchestrator_connection.log_trace("Running process.")

    job_name, udtraek_id = parse_queue_element(queue_element)
    orchestrator_connection.log_trace(f"Spool job {job_name}, window {udtraek_id}.")

    session = get_sap_session(connection_index=0, session_index=0)
    wait_ready(session)

    open_spool_overview(session)
    spool_session = wait_for_session(connection_index=0, session_index=1)
    wait_ready(spool_session)

    row = wait_for_spool_job(spool_session, job_name)
    select_spool_job(spool_session, row)
    export_spool_as_text(spool_session)
    file_path = get_exported_file_path(spool_session)
    orchestrator_connection.log_trace(f"Spool exported to: {file_path}")

    counts = mssql_load.load_spool_file(
        orchestrator_connection,
        file_path=file_path,
        spool_job=job_name,
        udtraek_id=udtraek_id,
    )
    orchestrator_connection.log_trace(
        "Merged into CJI3: "
        + ", ".join(f"{key}={value}" for key, value in counts.items())
    )

    if counts.get("Konverteringsadvarsler"):
        orchestrator_connection.log_info(
            f"{counts['Konverteringsadvarsler']} row(s) had a value that could not be "
            "converted to date/number. Check the spool column layout."
        )

    delete_export(orchestrator_connection, file_path)


def delete_export(orchestrator_connection, file_path: str) -> None:
    """
    Remove the exported file once its rows are safely committed.

    Only after load_spool_file has returned, which is after its COMMIT - so the file is
    never removed on the strength of a load that later rolled back.

    Worth doing: these run 9-39 MB each and SAP writes them to the user's Documents
    folder, so a full backfill would otherwise leave a couple of gigabytes behind.

    A failure here is logged and ignored. The rows are already in, and a leftover file
    is untidy rather than wrong - it must never turn a successful load into a failed
    queue element.
    """
    try:
        os.remove(file_path)
        orchestrator_connection.log_trace(f"Deleted export {file_path}")
    except OSError as error:
        orchestrator_connection.log_info(
            f"Could not delete the export {file_path}: {error}. The rows loaded fine; "
            "the file just needs clearing up."
        )


def parse_queue_element(queue_element: QueueElement) -> tuple[str, int | None]:
    """
    Read the spool job name and window id from a queue element.

    The dispatcher writes {"SpoolJob": ..., "UdtraekId": ...}. Elements created
    before the work list existed carry the bare prtxt as data, and are still
    accepted - they just merge without closing a window.
    """
    data = (queue_element.data or "").strip()
    if not data:
        raise BusinessError("Queue element has no data; cannot tell which spool job to fetch.")

    if not data.startswith("{"):
        return data, None

    payload = json.loads(data)
    job_name = payload.get("SpoolJob")
    if not job_name:
        raise BusinessError(f"Queue element data has no SpoolJob: {data}")

    return job_name, payload.get("UdtraekId")



def open_spool_overview(session) -> None:
    """Navigate to own spool jobs via System → Egne spooljobs (no SP01 auth needed).
    menu[4]/menu[8] is the recorded path — index depends on the active screen,
    so verify this if the robot is on a different screen when called.
    """
    session.findById("wnd[0]/mbar/menu[4]/menu[8]").select()
    wait_ready(session)


def wait_for_spool_job(session, job_name: str, timeout_s: int | None = None, poll_interval_s: int | None = None) -> int:
    """
    Poll the spool overview until a spool job whose title contains job_name appears.
    Returns the row number of the matching job.

    Raises BusinessError rather than TimeoutError on giving up. That is deliberate:
    queue_framework only keeps the queue loop alive for BusinessError, so a
    TimeoutError here would abandon every remaining queue element in the run - and
    when the dispatcher has queued ten windows for the night, one slow spool job
    would cost the other nine. The window itself is not lost: it stays IGang in
    dbo.CJI3_Udtraek until the stale sweep in usp_CJI3_ReserverUdtraek
    puts it back to Afventer, and the next dispatcher run submits it again.
    """
    timeout_s = config.SPOOL_TIMEOUT_S if timeout_s is None else timeout_s
    poll_interval_s = config.SPOOL_POLL_INTERVAL_S if poll_interval_s is None else poll_interval_s

    deadline = time.time() + timeout_s
    com_errors = 0          # consecutive; reset on any successful read
    total_com_errors = 0

    while True:
        try:
            row = _find_spool_job_row(session, job_name)
            com_errors = 0
        except pywintypes.com_error as error:
            # SAP answers E_PENDING while it is busy with a server round-trip, which is
            # a "not yet", not a failure. Seen in practice after ~22 minutes of polling,
            # where it aborted the whole run. Only a long unbroken streak is fatal.
            com_errors += 1
            total_com_errors += 1
            if com_errors >= config.SPOOL_MAX_CONSECUTIVE_COM_ERRORS:
                raise BusinessError(
                    f"SAP GUI has refused to be read {com_errors} times in a row while "
                    f"waiting for spool job '{job_name}': {error}. The session looks "
                    "wedged rather than busy."
                ) from error
            row = None

        if row is not None:
            return row

        if time.time() >= deadline:
            raise BusinessError(
                f"Spool job matching '{job_name}' was not ready within {timeout_s}s "
                f"({total_com_errors} transient SAP read error(s) along the way). Either "
                "SAP is still generating it, or the row is not on the visible page of "
                "the spool overview."
            )

        session.findById("wnd[0]/tbar[1]/btn[45]").press()  # Opdater (Ctrl+Shift+F9)
        wait_ready(session)
        # Sleep AFTER the refresh, not before: it paces the polling and gives SAP time
        # to finish repainting before the next read, which is when E_PENDING appears.
        time.sleep(poll_interval_s)


def _find_spool_job_row(session, job_name: str) -> int | None:
    """Scan the detected title column for the first row whose title contains job_name.
    Returns None if not found yet or if the job is still generating (status '+').
    """
    labels = _get_spool_label_cells(session)
    title_x = _get_spool_column_x(labels, "Titel")
    status_x = _get_spool_column_x(labels, "Status")
    labels_by_position = {(x, y): text for x, y, text in labels}

    for row in _get_spool_row_numbers(session):
        title = labels_by_position.get((title_x, row), "")
        if job_name in title:
            status = labels_by_position.get((status_x, row))
            if status is None:
                raise RuntimeError(f"Found spool job at row {row}, but no Status cell exists on that row.")
            if status == "+":
                return None  # Still generating
            return row
    return None


def _get_spool_column_x(labels: list[tuple[int, int, str]], header_text: str) -> int:
    """Return the x-coordinate for an exact spool overview column header."""
    matches = [
        (x, y)
        for x, y, text in labels
        if text.casefold() == header_text.casefold()
    ]
    if not matches:
        visible_headers = ", ".join(
            f"{text!r} at x={x},y={y}"
            for x, y, text in labels
            if y == 1 and text
        )
        raise RuntimeError(
            f"Could not find spool column header {header_text!r}. "
            f"Visible headers: {visible_headers or 'none'}"
        )
    x, _ = min(matches, key=lambda position: position[1])
    return x


def _get_spool_label_cells(session) -> list[tuple[int, int, str]]:
    """Return all visible SAP label cells as (x, y, text)."""
    usr = session.findById("wnd[0]/usr")
    cells = []
    for index in range(usr.Children.Count):
        control = usr.Children(index)
        match = _SAP_LABEL_ID_PATTERN.search(control.Id)
        if not match:
            continue
        try:
            text = control.Text.strip()
        except Exception:  # SAP GUI scripting can raise COM exceptions for missing text.
            text = ""
        cells.append((int(match.group(1)), int(match.group(2)), text))
    return cells


def _get_spool_row_numbers(session) -> list[int]:
    """Return visible selectable spool rows, based on the row checkboxes."""
    usr = session.findById("wnd[0]/usr")
    rows = []
    for index in range(usr.Children.Count):
        control = usr.Children(index)
        match = _SAP_CHECKBOX_ID_PATTERN.search(control.Id)
        if match:
            rows.append(int(match.group(2)))
    return sorted(set(rows))


def select_spool_job(session, row: int) -> None:
    """
    Select exactly one spool row, clearing every other selection first.

    The clearing is the important half. This robot processes several queue elements per
    run and the spool overview session survives between them, so a row ticked for one
    element was still ticked for the next. SAP then exported BOTH spool jobs and reported
    'Filer arkiveret i directory ...' - plural, and with no filename - instead of
    'Fil <navn> gemt i directory <sti>'. get_exported_file_path could not parse that, and
    the ValueError took down the whole run.

    Seen for real: element 2 exported P020000400765.TXT and loaded 32,273 rows, then
    element 3's export rewrote that same file alongside its own P020000401621.TXT.
    """
    target_id = _get_spool_row_checkbox_id(session, row)

    for checkbox_id in _all_spool_checkbox_ids(session):
        if checkbox_id == target_id:
            continue
        try:
            checkbox = session.findById(checkbox_id)
            if checkbox.Selected:
                checkbox.Selected = False
        except Exception:  # pylint: disable=broad-exception-caught
            # A row that has scrolled away is not selectable and cannot be exported
            # either, so it does not matter.
            continue

    session.findById(target_id).Selected = True


def _all_spool_checkbox_ids(session) -> list[str]:
    """Every row-selection checkbox id currently on the spool overview."""
    usr = session.findById("wnd[0]/usr")
    ids = []
    for index in range(usr.Children.Count):
        control = usr.Children(index)
        try:
            if _SAP_CHECKBOX_ID_PATTERN.search(control.Id):
                ids.append(control.Id)
        except Exception:  # pylint: disable=broad-exception-caught
            continue
    return ids


def _get_spool_row_checkbox_id(session, row: int) -> str:
    """Return the checkbox id for a visible spool overview row."""
    usr = session.findById("wnd[0]/usr")
    checkboxes = []
    for index in range(usr.Children.Count):
        control = usr.Children(index)
        match = _SAP_CHECKBOX_ID_PATTERN.search(control.Id)
        if match and int(match.group(2)) == row:
            checkboxes.append((int(match.group(1)), control.Id))
    if not checkboxes:
        raise RuntimeError(f"Could not find selection checkbox for spool row {row}.")
    _, checkbox_id = min(checkboxes, key=lambda item: item[0])
    return checkbox_id


def export_spool_as_text(session) -> None:
    """
    Export the selected spool job via Spooljob > Videresend > 'Eksporter som tekst'.

    menu[1], NOT menu[3]. menu[3] is 'Tekst med tabulator', which was used originally
    and produces a file that cannot be parsed reliably: it puts a tab at each print
    column boundary, so a cell that does not fill its column changes the tab count and
    shifts every later value on the row. On one real 17,014-row export that corrupted
    75.7% of rows. menu[1] writes pipe-delimited fixed-width text instead, which parsed
    with zero bad values on the same data. See spool_to_sql for the detail.

    Despite the ellipsis in the menu entry it opens no dialog - it writes the file
    straight away and reports the path in the status bar, exactly as menu[3] did, so
    get_exported_file_path is unchanged.
    """
    session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[1]").select()
    wait_ready(session)


def get_exported_file_path(session) -> str:
    """
    Parse the exported file path from the status bar.
    Expected format: 'Fil <filename> gemt i directory <directory>'
    """
    status_text = session.findById("wnd[0]/sbar/pane[0]").Text
    match = re.search(r"Fil\s+(\S+)\s+gemt i directory\s+(.+)", status_text)
    if not match:
        raise ValueError(f"Unexpected status bar message: {status_text!r}")
    filename = match.group(1)
    directory = match.group(2).strip()
    if directory.upper().endswith(r'\SAP\SAP'):
        directory = directory[:-3] + 'SAP GUI'
    return os.path.join(directory, filename)


def wait_for_session(connection_index: int = 0, session_index: int = 1, timeout_s: int = 15):
    """Poll until the given SAP session index exists (e.g. after a new window opens)."""
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    try:
        connection = application.Children(connection_index)
    except Exception:
        connection = application.Connections(connection_index)
    deadline = time.time() + timeout_s
    while True:
        try:
            return connection.Children(session_index)
        except Exception:
            if time.time() >= deadline:
                raise TimeoutError(f"SAP session {session_index} did not appear within {timeout_s}s")
            time.sleep(0.5)


def get_sap_session(connection_index: int = 0, session_index: int = 0):
    """
    Get an active SAP GUI Scripting session.
    Assumes SAP GUI is open and you're logged in.
    """
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine

    try:
        connection = application.Children(connection_index)
    except Exception:
        connection = application.Connections(connection_index)

    session = connection.Children(session_index)
    return session


def wait_ready(session, timeout_s: int = 30) -> None:
    """Wait until SAP session is not busy."""
    start = time.time()
    while getattr(session, "Busy", False):
        if time.time() - start > timeout_s:
            raise TimeoutError("SAP session stayed busy too long.")
        time.sleep(0.1)


