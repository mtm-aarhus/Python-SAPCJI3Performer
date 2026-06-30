"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

from robot_framework.process import process
from robot_framework import config
from robot_framework import reset
from OpenOrchestrator.database.queues import QueueStatus
from robot_framework.exceptions import handle_error, BusinessError

import os

# pylint: disable-next=unused-argum
orchestrator_connection = OrchestratorConnection(
    "SAPCJI3Performer",
    os.getenv("OpenOrchestratorSQL"),
    os.getenv("OpenOrchestratorKey"),
    None,
    None,
    None
)

reset.reset(orchestrator_connection)

queue_element = orchestrator_connection.get_next_queue_element(config.QUEUE_NAME)

try:
    process(orchestrator_connection, queue_element)
    orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.DONE)
except BusinessError as error:
    handle_error("Business Error", error, queue_element, orchestrator_connection)