"""This module is the primary module of the robot framework. It collects the functionality of the rest of the framework."""

# This module is not meant to exist next to linear_framework.py in production:
# pylint: disable=duplicate-code

import sys

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueStatus

from robot_framework import initialize
from robot_framework import reset
from robot_framework.exceptions import handle_error, BusinessError, log_exception
from robot_framework import process
from robot_framework import config


def main():
    """
    The entry point for the framework. Should be called as the first thing when running the robot.

    Failures are separated into two kinds, because they need opposite handling:

    SETUP (reset/initialize) - SAP will not launch, Opus login broken. Nothing can work,
        so this fails the robot. Queue elements are left untouched: they stay NEW and the
        next run picks them up.

    ELEMENT - one queue element. Retried up to config.QUEUE_ATTEMPTS times with a full
        reset between attempts, then marked FAILED, and the run CONTINUES with the rest
        of the queue. These do not fail the robot: failing it puts the trigger into a
        failed state, and a queue trigger only fires while it is IDLE, so one bad element
        would stop the queue being served until somebody reactivated it by hand.

    That separation is the point. A ValueError from the export path once took down a
    whole run: three elements, the first timed out harmlessly, the second loaded 32,273
    rows, and the third raised ValueError, which escaped the BusinessError handler and
    killed the robot and the trigger with it.
    """
    orchestrator_connection = OrchestratorConnection.create_connection_from_args()
    sys.excepthook = log_exception(orchestrator_connection)

    orchestrator_connection.log_trace("Robot Framework started.")
    initialize.initialize(orchestrator_connection)

    queue_element = None
    error_count = 0        # element failures - do NOT fail the robot
    setup_error_count = 0  # reset/initialize failures - DO fail the robot
    task_count = 0

    # Retry loop
    for _ in range(config.MAX_RETRY_COUNT):
        try:
            reset.reset(orchestrator_connection)

            # Queue loop
            while task_count < config.MAX_TASK_COUNT:
                task_count += 1
                queue_element = orchestrator_connection.get_next_queue_element(config.QUEUE_NAME)

                if not queue_element:
                    orchestrator_connection.log_info("Queue empty.")
                    break  # Break queue loop

                try:
                    for attempt in range(1, config.QUEUE_ATTEMPTS + 1):
                        try:
                            process.process(orchestrator_connection, queue_element)
                            break

                        except BusinessError:
                            # A business error will not fix itself on a retry, and here a
                            # retry is expensive: a spool job that was never generated
                            # costs SPOOL_TIMEOUT_S of polling plus a full SAP relaunch
                            # per attempt. Fail this element now and move on.
                            raise

                        # pylint: disable-next = broad-exception-caught
                        except Exception as e:
                            orchestrator_connection.log_trace(f"Attempt {attempt} failed for current queue element: {e}")
                            if attempt < config.QUEUE_ATTEMPTS:
                                orchestrator_connection.log_trace("Retrying queue element.")
                                reset.reset(orchestrator_connection)
                            else:
                                orchestrator_connection.log_trace(f"Queue element failed after {attempt} attempts.")
                                raise
                    orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.DONE)

                except BusinessError as error:
                    handle_error("Business Error", error, queue_element, orchestrator_connection)

                # pylint: disable-next = broad-exception-caught
                except Exception as error:
                    # Isoler fejl pr. koeelement: markeer FAILED og fortsaet med resten af koeen
                    error_count += 1
                    handle_error(f"Process Error #{error_count}", error, queue_element, orchestrator_connection)
                    reset.reset(orchestrator_connection)

            break  # Break retry loop

        # We actually want to catch all exceptions possible here.
        # pylint: disable-next = broad-exception-caught
        except Exception as error:
            setup_error_count += 1
            handle_error(f"Setup Error #{setup_error_count}", error, queue_element, orchestrator_connection)

    reset.clean_up(orchestrator_connection)
    reset.close_all(orchestrator_connection)
    reset.kill_all(orchestrator_connection)

    if error_count:
        orchestrator_connection.log_info(
            f"{error_count} queue element(s) failed and were marked FAILED. The run itself "
            "is not failed, so the queue trigger stays active."
        )

    # Only setup failures fail the robot - see the docstring.
    if config.FAIL_ROBOT_ON_TOO_MANY_ERRORS and setup_error_count >= config.MAX_RETRY_COUNT:
        raise RuntimeError("Process failed too many times.")
