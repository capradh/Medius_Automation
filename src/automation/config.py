"""Configuration loading.

Databases are discovered from environment variables following the pattern
DB_<NAME>_USER, DB_<NAME>_PASSWORD, DB_<NAME>_DSN and optional DB_<NAME>_ENV.
Adding a new database therefore needs only new lines in .env.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Mapping

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parents[2]
load_dotenv(PROJECT_ROOT / ".env")

_KEY = re.compile(r"^DB_(?P<name>[A-Z0-9_]+?)_(?P<field>USER|PASSWORD|DSN|ENV)$")
_REQUIRED = ("USER", "PASSWORD", "DSN")


class ConfigError(Exception):
    """Raised for missing or invalid configuration."""


@dataclass(frozen=True)
class DatabaseConfig:
    name: str
    user: str
    dsn: str
    password: str = field(repr=False)  # never printed in logs/tracebacks
    env: str = "dev"

    @property
    def is_prod(self) -> bool:
        return self.env.lower() in ("prod", "production")


@dataclass(frozen=True)
class Settings:
    log_level: str
    default_db: str | None
    oracle_client_lib_dir: str | None
    output_dir: Path
    log_dir: Path
    input_dir: Path


@dataclass(frozen=True)
class SmtpSettings:
    host: str | None
    port: int
    security: str  # starttls | ssl | none
    user: str | None
    password: str | None = field(repr=False)
    sender: str | None
    default_to: tuple[str, ...]
    timeout: int = 30


@dataclass(frozen=True)
class ApiSettings:
    base_url: str | None
    api_key: str | None = field(repr=False)
    key_header: str = "Authorization"
    key_prefix: str = "Bearer"
    timeout: int = 30
    max_retries: int = 3
    verify: bool | str = True


def discover_databases(environ: Mapping[str, str] | None = None) -> dict[str, DatabaseConfig]:
    """Return {NAME: DatabaseConfig} for every fully configured DB_<NAME>_* group."""
    environ = os.environ if environ is None else environ
    groups: dict[str, dict[str, str]] = {}
    for key, value in environ.items():
        m = _KEY.match(key.upper())
        if m and value.strip():
            groups.setdefault(m["name"], {})[m["field"]] = value.strip()

    result: dict[str, DatabaseConfig] = {}
    for name, f in groups.items():
        if all(k in f for k in _REQUIRED):
            result[name] = DatabaseConfig(
                name=name, user=f["USER"], password=f["PASSWORD"],
                dsn=f["DSN"], env=f.get("ENV", "dev"),
            )
    return result


def list_incomplete(environ: Mapping[str, str] | None = None) -> dict[str, list[str]]:
    """Databases that have some DB_<NAME>_* values but are missing required ones."""
    environ = os.environ if environ is None else environ
    groups: dict[str, set[str]] = {}
    for key, value in environ.items():
        m = _KEY.match(key.upper())
        if m and value.strip():
            groups.setdefault(m["name"], set()).add(m["field"])
    return {n: [k for k in _REQUIRED if k not in f]
            for n, f in groups.items() if any(k not in f for k in _REQUIRED)}


def get_database(name: str, environ: Mapping[str, str] | None = None) -> DatabaseConfig:
    dbs = discover_databases(environ)
    key = name.strip().upper()
    if key in dbs:
        return dbs[key]
    incomplete = list_incomplete(environ)
    if key in incomplete:
        raise ConfigError(
            f"Database '{key}' is missing: "
            + ", ".join(f"DB_{key}_{k}" for k in incomplete[key])
            + ". Check your .env file."
        )
    available = ", ".join(sorted(dbs)) or "(none configured)"
    raise ConfigError(f"Unknown database '{name}'. Configured databases: {available}")


def get_settings(environ: Mapping[str, str] | None = None) -> Settings:
    environ = os.environ if environ is None else environ
    return Settings(
        log_level=environ.get("LOG_LEVEL", "INFO").upper(),
        default_db=(environ.get("DEFAULT_DB") or "").strip().upper() or None,
        oracle_client_lib_dir=environ.get("ORACLE_CLIENT_LIB_DIR") or None,
        output_dir=PROJECT_ROOT / "output",
        log_dir=PROJECT_ROOT / "logs",
        input_dir=PROJECT_ROOT / "input",
    )


def split_addresses(value: str | None) -> list[str]:
    return [a.strip() for a in re.split(r"[;,]", value or "") if a.strip()]


def get_smtp_settings(environ: Mapping[str, str] | None = None) -> SmtpSettings:
    environ = os.environ if environ is None else environ
    user = (environ.get("SMTP_USER") or "").strip() or None
    security = (environ.get("SMTP_SECURITY") or "starttls").strip().lower()
    if security not in ("starttls", "ssl", "none"):
        raise ConfigError("SMTP_SECURITY must be one of: starttls, ssl, none")
    return SmtpSettings(
        host=(environ.get("SMTP_HOST") or "").strip() or None,
        port=int(environ.get("SMTP_PORT") or 587),
        security=security,
        user=user,
        password=environ.get("SMTP_PASSWORD") or None,
        sender=(environ.get("EMAIL_FROM") or "").strip() or user,
        default_to=tuple(split_addresses(environ.get("EMAIL_TO"))),
    )


def get_api_settings(environ: Mapping[str, str] | None = None) -> ApiSettings:
    environ = os.environ if environ is None else environ
    verify_raw = (environ.get("API_VERIFY_SSL") or "true").strip()
    verify: bool | str = verify_raw
    if verify_raw.lower() in ("true", "1", "yes"):
        verify = True
    elif verify_raw.lower() in ("false", "0", "no"):
        verify = False
    return ApiSettings(
        base_url=(environ.get("API_BASE_URL") or "").strip() or None,
        api_key=(environ.get("API_KEY") or "").strip() or None,
        key_header=(environ.get("API_KEY_HEADER") or "Authorization").strip(),
        key_prefix=(environ.get("API_KEY_PREFIX") if environ.get("API_KEY_PREFIX") is not None else "Bearer").strip(),
        timeout=int(environ.get("API_TIMEOUT") or 30),
        max_retries=int(environ.get("API_MAX_RETRIES") or 3),
        verify=verify,
    )
