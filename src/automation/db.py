"""Oracle connection helpers. Every function takes the database name (e.g. 'corsair')."""
from __future__ import annotations

import logging
from contextlib import contextmanager
from typing import Any, Iterator

import oracledb

from .config import get_database, get_settings

log = logging.getLogger(__name__)
_client_initialised = False


def _init_thick_mode_if_requested() -> None:
    """Switch to thick mode only if ORACLE_CLIENT_LIB_DIR is set (default is thin mode)."""
    global _client_initialised
    lib_dir = get_settings().oracle_client_lib_dir
    if lib_dir and not _client_initialised:
        oracledb.init_oracle_client(lib_dir=lib_dir)
        _client_initialised = True
        log.info("Oracle thick mode enabled (%s)", lib_dir)


@contextmanager
def get_connection(db_name: str, timeout: int = 10) -> Iterator[oracledb.Connection]:
    cfg = get_database(db_name)
    _init_thick_mode_if_requested()
    log.debug("Connecting to %s (%s) as %s", cfg.name, cfg.env, cfg.user)
    conn = oracledb.connect(
        user=cfg.user, password=cfg.password, dsn=cfg.dsn, tcp_connect_timeout=timeout
    )
    try:
        yield conn
    finally:
        conn.close()


def fetch_all(db_name: str, sql: str, params: dict[str, Any] | None = None,
              max_rows: int | None = None) -> list[dict[str, Any]]:
    """Run a SELECT and return rows as dicts. Use bind variables (:name) for values."""
    with get_connection(db_name) as conn, conn.cursor() as cur:
        cur.execute(sql, params or {})
        cols = [c[0] for c in cur.description]
        rows = cur.fetchmany(max_rows) if max_rows else cur.fetchall()
        return [dict(zip(cols, r)) for r in rows]


def execute_write(db_name: str, sql: str, params: dict[str, Any] | None = None,
                  dry_run: bool = True) -> int:
    """Run INSERT/UPDATE/DELETE. With dry_run=True the change is rolled back.

    Returns the number of rows affected (also in dry-run, so you can review it first).
    """
    with get_connection(db_name) as conn, conn.cursor() as cur:
        cur.execute(sql, params or {})
        affected = cur.rowcount
        if dry_run:
            conn.rollback()
            log.info("[DRY RUN] %d row(s) would be affected on %s. Rolled back.", affected, db_name)
        else:
            conn.commit()
            log.info("Committed: %d row(s) affected on %s.", affected, db_name)
        return affected
