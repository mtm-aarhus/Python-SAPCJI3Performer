"""This module contains configuration constants used across the framework"""

# The number of times the robot retries on an error before terminating.
MAX_RETRY_COUNT = 1

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
MAX_TASK_COUNT = 100

# ----------------------


# SAP spool
# ----------------------

# How long to wait for SAP to finish generating a spool job. Generation has been
# observed to take up to 10 minutes for a week of data, so this is set well above
# that - the robot polls and returns as soon as the job is ready, so a high
# ceiling costs nothing on a normal run.
SPOOL_TIMEOUT_S = 1800

# How often to refresh the spool overview while waiting.
SPOOL_POLL_INTERVAL_S = 15

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
