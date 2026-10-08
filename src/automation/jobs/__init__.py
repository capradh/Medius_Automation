"""Each module here (not starting with '_') is one automation job.

A job module provides:
    run(args) -> JobResult | None          (required)
    add_arguments(parser)                  (optional: job-specific command-line options)
    REQUIRES_DB = True                     (optional: job needs --db)
    WRITES = True                          (optional: job can modify data; prod asks for confirmation)
"""
