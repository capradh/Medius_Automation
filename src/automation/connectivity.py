"""Connectivity diagnostics: TCP reachability, login, and basic DB info."""
from __future__ import annotations

import re
import socket
import time
from dataclasses import dataclass

import oracledb

from .config import DatabaseConfig
from .db import get_connection

_EZ = re.compile(r"^(?://)?(?P<host>[^:/\s]+)(?::(?P<port>\d+))?(?:/.*)?$")

_HINTS = {
    "ORA-01017": "Wrong username/password (or account locked).",
    "ORA-28000": "Account is locked. Ask the DBA to unlock it.",
    "ORA-28001": "Password expired. Reset it with your DBA.",
    "ORA-12541": "No listener at that host/port. Check host, port, and VPN.",
    "ORA-12514": "Listener doesn't know that service name. Check the service name in the DSN.",
    "ORA-12154": "Could not resolve the connect identifier. Check the DSN format.",
    "ORA-12170": "Connection timed out. Likely firewall/VPN/IP allowlist.",
    "ORA-12545": "Host not found. Check the hostname and your DNS/VPN.",
    "DPY-6005": "Could not connect. Check network/VPN, host, port, and service name.",
    "DPY-6001": "Listener rejected the service name. Check the service name in the DSN.",
    "DPY-3010": "Server version needs thick mode. Set ORACLE_CLIENT_LIB_DIR.",
    "DPY-4011": "Connection closed by server. Often network/firewall related.",
}


def parse_host_port(dsn: str) -> tuple[str, int] | None:
    """Extract host/port from an Easy Connect string. Returns None for TNS descriptors/aliases."""
    if "(" in dsn:  # full TNS descriptor
        h = re.search(r"HOST\s*=\s*([^)\s]+)", dsn, re.I)
        p = re.search(r"PORT\s*=\s*(\d+)", dsn, re.I)
        return (h[1], int(p[1])) if h else None
    m = _EZ.match(dsn.strip())
    if not m or "/" not in dsn and ":" not in dsn:
        return None
    return m["host"], int(m["port"] or 1521)


def hint_for(error: Exception) -> str:
    text = str(error)
    for code, hint in _HINTS.items():
        if code in text:
            return hint
    return "See the full error above."


@dataclass
class CheckResult:
    name: str
    env: str
    dsn: str
    tcp_ok: bool | None = None
    login_ok: bool = False
    tcp_ms: float | None = None
    login_ms: float | None = None
    detail: str = ""
    error: str = ""
    hint: str = ""

    @property
    def ok(self) -> bool:
        return self.login_ok


def check_database(cfg: DatabaseConfig, timeout: int = 10) -> CheckResult:
    res = CheckResult(name=cfg.name, env=cfg.env, dsn=cfg.dsn)

    hp = parse_host_port(cfg.dsn)
    if hp:
        t0 = time.perf_counter()
        try:
            with socket.create_connection(hp, timeout=min(timeout, 5)):
                res.tcp_ok = True
        except OSError as e:
            res.tcp_ok = False
            res.error = f"TCP connect to {hp[0]}:{hp[1]} failed: {e}"
            res.hint = "Network issue: check VPN, firewall, IP allowlist, host and port."
            return res
        finally:
            res.tcp_ms = (time.perf_counter() - t0) * 1000

    t0 = time.perf_counter()
    try:
        with get_connection(cfg.name, timeout=timeout) as conn, conn.cursor() as cur:
            cur.execute(
                "SELECT USER, SYS_CONTEXT('USERENV','DB_NAME'), "
                "SYS_CONTEXT('USERENV','SERVICE_NAME') FROM dual"
            )
            user, db_name, service = cur.fetchone()
            res.login_ok = True
            res.detail = (f"user={user} db={db_name} service={service} "
                          f"version={conn.version} mode={'thin' if conn.thin else 'thick'}")
    except (oracledb.Error, Exception) as e:  # noqa: BLE001 - diagnostics must never crash
        res.error = str(e).strip().splitlines()[0] if str(e).strip() else repr(e)
        res.hint = hint_for(e)
    finally:
        res.login_ms = (time.perf_counter() - t0) * 1000
    return res
