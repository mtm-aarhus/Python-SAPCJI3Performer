"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement

import json
import os
import re
import time

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
    export_spool_as_tabtext(spool_session)
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
    while True:
        row = _find_spool_job_row(session, job_name)
        if row is not None:
            return row
        if time.time() >= deadline:
            raise BusinessError(
                f"Spool job matching '{job_name}' was not ready within {timeout_s}s. "
                "Either SAP is still generating it, or the row is not on the visible "
                "page of the spool overview."
            )
        time.sleep(poll_interval_s)
        session.findById("wnd[0]/tbar[1]/btn[45]").press()  # Opdater (Ctrl+Shift+F9)
        wait_ready(session)


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
    """Set the selection checkbox for the given spool overview row."""
    checkbox_id = _get_spool_row_checkbox_id(session, row)
    session.findById(checkbox_id).Selected = True


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


def export_spool_as_tabtext(session) -> None:
    """Export the selected spool job as tab-separated text via Spooljob > Videresend > Tekst med tabulatorer."""
    session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[3]").select()
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


