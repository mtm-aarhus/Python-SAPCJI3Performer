"""This module contains configuration constants used across the framework"""

# The number of times the robot retries on an error before terminating.
# Now counts SETUP failures only (reset/initialize): a queue element that fails no
# longer fails the robot, because that would put the trigger into a failed state and a
# queue trigger only fires while it is IDLE.
MAX_RETRY_COUNT = 1

# How many times one queue element is retried before it is marked FAILED and the run
# moves on. A full reset - SAP relaunch and Opus login - happens between attempts, so
# this is the recovery path for a wedged or lost SAP session.
#
# It does NOT apply to BusinessError, which is raised straight through: a spool job that
# was never generated would otherwise cost SPOOL_TIMEOUT_S of polling plus a relaunch on
# every attempt, and no amount of retrying will make it appear.
QUEUE_ATTEMPTS = 2

# Whether the robot should be marked as failed if MAX_RETRY_COUNT is reached.
FAIL_ROBOT_ON_TOO_MANY_ERRORS = True

# Error screenshot config
SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
SMTP_PORT = 25
SCREENSHOT_SENDER = "sapcji3@aarhus.dk"

# Constant/Credential names
ERROR_EMAIL = "Error Email"


# Queue specific configs
# ----------------------

# The name of the job queue (if any)
QUEUE_NAME = "SAPCJI3"

# The limit on how many queue elements to process
MAX_TASK_COUNT = 5

# ----------------------


# SAP spool
# ----------------------

# How long to wait for SAP to finish generating a spool job.
#
# Measured from SM37: report RKPEP003 finishes a full week in 320-445 seconds, so ~7
# minutes is normal and 30 gives ample headroom. An earlier 22-minute wait was NOT slow
# generation - the background job had finished in 7 minutes and produced no spool at
# all, so the robot was polling for something that would never appear. Since that is
# the case this ceiling actually bounds, keep it modest: during an 85-window backfill,
# every window with no data costs this long before it gives up.
SPOOL_TIMEOUT_S = 1800          # 30 minutes

# How often to refresh the spool overview while waiting. Doubles as the settle pause
# after each refresh, before the screen is read again.
SPOOL_POLL_INTERVAL_S = 15

# SAP GUI scripting raises E_PENDING (0x8000000A, "the data required is not yet
# available") when a control is read while SAP is mid round-trip. It means "not yet",
# not "broken", so those are retried rather than failing the run. This caps how many
# CONSECUTIVE ones are tolerated, so a genuinely wedged session still gives up:
# 20 x SPOOL_POLL_INTERVAL_S = 5 minutes of nothing but errors.
SPOOL_MAX_CONSECUTIVE_COM_ERRORS = 20

# ----------------------


# MSSQL (BI_Oekonomi.dbo.CJI3)
# ----------------------

# Name of the OpenOrchestrator constant holding the SQL server host.
SQL_SERVER_CONSTANT = "SqlServer"

SQL_DATABASE = "BI_Oekonomi"

# Not the legacy "SQL Server" driver: that one supports neither the DATE/TIME
# types nor pyodbc's fast_executemany.
SQL_DRIVER = "ODBC Driver 17 for SQL Server"

# Rows per executemany batch when filling the staging table.
STAGE_CHUNK_SIZE = 1000

# ----------------------
