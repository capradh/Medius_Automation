r"""Update Medius user Roles and Authorization Groups from an Excel access matrix.

Excel layout: CCGID | Username | First Name | Last Name | Email | Company Code | <role columns> | ;<auth group columns>
Cell x/X = assign, blank = ignore, and NA or N/A = remove.
Only users with an @cma-cgm.com email get payloads; users with any other domain are
left out of the JSON payloads and reported as errors (summary + Excel report).

python scripts\run_job.py medius_user_access --file input\users.xlsx --env QA
python scripts\run_job.py medius_user_access --file input\users.xlsx --env QA --execute --notify
python scripts\run_job.py medius_user_access --file input\users.xlsx --env PROD --execute --confirm-production

.\.venv\Scripts\python.exe scripts\run_job.py medius_user_access --pick --env QA
.\.venv\Scripts\python.exe scripts\run_job.py medius_user_access --pick --env QA --execute
"""

from __future__ import annotations

import json
import logging
import os
import random
import re
import time
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from openpyxl.utils import get_column_letter

from ..config import get_settings
from ..results import JobResult
from ..io_files import pick_files, write_excel

log = logging.getLogger(__name__)

WRITES = True  # --execute sends data to Medius; default is dry run

# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #
TIMEZONE, LANGUAGE, FORMAT = "UTC", "en-US", "en-US"
CONNECT_TIMEOUT, READ_TIMEOUT = 10.0, 60.0
MAX_RETRIES, BACKOFF_BASE, BACKOFF_MAX = 4, 1.0, 30.0
RETRY_STATUSES = frozenset({429, 502, 503, 504})
TOKEN_SKEW = 60.0
HEADER_SCAN_ROWS = 50
MIN_HEADER_SCORE = 3
MAX_SUMMARY_LINES = 50

IGNORED_COLUMNS = (
    "Comments",
    "Comment",
    "Remarks",
    "Remark",
    "Notes",
    "Note",
    "Sr No",
    "S.No",
    "Sl No",
    "No",
    "#",
    "Serial No",
    "Department",
    "Manager",
    "Requester",
    "Ticket",
    "Ticket No",
)

USER_COLUMN_ALIASES: Mapping[str, Tuple[str, ...]] = {
    "external_id": ("CCGID", "CCG ID", "CCG Id"),
    "user_name": ("Username", "User Name", "User Id", "Login"),
    "first_name": ("First Name", "Firstname", "Given Name"),
    "last_name": ("Last Name", "Lastname", "Surname", "Family Name"),
    "email": ("Email", "E-mail", "Email Address", "Mail"),
    "company": ("Company Code", "CompanyCode", "Company Id"),
}
IDENTITY_FIELDS = ("external_id", "user_name", "first_name", "last_name", "email")
FIELD_LABELS = {k: v[0] for k, v in USER_COLUMN_ALIASES.items()}

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
ALLOWED_EMAIL_DOMAIN = "cma-cgm.com"  # exact match, case-insensitive
ASSIGN_MARKS = frozenset({"X", "\u00d7"})
IGNORE_MARKS = frozenset({"", "N/A", "NA", "N.A.", "N.A", "N\\A", "-", "--", "#N/A"})
REMOVAL_MARKS = frozenset({"NA", "N/A"})
_CYRILLIC_TO_LATIN = str.maketrans({"\u0425": "X", "\u0445": "X"})
_INVISIBLE_RE = re.compile("[\u200b\u200c\u200d\u2060\ufeff]")
_FLOAT_ID_RE = re.compile(r"^(\d+)\.0+$")
ALLOWED_SUFFIXES = (".xlsx", ".xlsm", ".xls")

S_OK, S_FAIL, S_SKIP, S_NA, S_DRY = "SUCCESS", "FAILED", "SKIPPED", "N/A", "DRY_RUN"
S_REJ = "REJECTED"


# --------------------------------------------------------------------------- #
# Data classes
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class EnvConfig:
    name: str
    token_url: str
    roles_url: str
    auth_groups_url: str
    identity_provider: str
    client_id: str = field(default="", repr=False)
    client_secret: str = field(default="", repr=False)


@dataclass(frozen=True)
class Issue:
    severity: str  # "ERROR" | "WARNING"
    message: str
    excel_row: Optional[int] = None

    def __str__(self) -> str:
        where = f"Row {self.excel_row}: " if self.excel_row else ""
        return f"[{self.severity}] {where}{self.message}"


@dataclass
class AccessRow:
    excel_row: int
    company: str
    roles: List[str] = field(default_factory=list)
    groups: List[str] = field(default_factory=list)
    roles_to_remove: List[str] = field(default_factory=list)
    groups_to_remove: List[str] = field(default_factory=list)


@dataclass
class UserRecord:
    user_name: str
    external_id: str
    first_name: str
    last_name: str
    email: str
    first_row: int
    rows: List[AccessRow] = field(default_factory=list)


@dataclass(frozen=True)
class RejectedUser:
    """User left out of the payloads because the email domain is not allowed."""

    user_name: str
    external_id: str
    email: str
    excel_rows: Tuple[int, ...]

    @property
    def reason(self) -> str:
        return (
            f"Email '{self.email}' is not an @{ALLOWED_EMAIL_DOMAIN} address - "
            "excluded from payloads"
        )

    def __str__(self) -> str:
        rows = ", ".join(str(r) for r in self.excel_rows)
        return f"[ERROR] Row {rows}: {self.user_name} - {self.reason}"


@dataclass(frozen=True)
class ColumnMap:
    user_pos: Mapping[str, int]
    roles: Mapping[str, Tuple[int, ...]]
    groups: Mapping[str, Tuple[int, ...]]


@dataclass(frozen=True)
class ApiResult:
    ok: bool
    status_code: Optional[int]
    message: str = ""
    import_link: str = ""


@dataclass
class UserResult:
    user_name: str
    external_id: str
    environment: str
    roles_status: str = S_NA
    roles_http: Optional[int] = None
    roles_message: str = ""
    roles_link: str = ""
    groups_status: str = S_NA
    groups_http: Optional[int] = None
    groups_message: str = ""
    groups_link: str = ""
    timestamp: str = ""


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #
def now_utc() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")


def _norm(text: str) -> str:
    return re.sub(r"[\s_\-]+", "", text).lower()


def _col_letters(positions: Sequence[int]) -> str:
    return ", ".join(get_column_letter(p + 1) for p in positions)


def _clean_text(value: Any) -> str:
    text = _INVISIBLE_RE.sub("", str(value).replace("\u00a0", " "))
    return re.sub(r"\s+", " ", text).strip()


def _clean_header(value: Any) -> str:
    return unicodedata.normalize("NFKC", _clean_text(value))


def _clean_id(value: str) -> str:
    """'992025.0' -> '992025' (IDs saved as numbers); text IDs untouched."""
    match = _FLOAT_ID_RE.match(value)
    return match.group(1) if match else value


def _mark(value: str) -> str:
    return (
        unicodedata.normalize("NFKC", value)
        .upper()
        .translate(_CYRILLIC_TO_LATIN)
        .strip()
    )


_ALIAS_TO_FIELD: Mapping[str, str] = {
    _norm(a): f for f, aliases in USER_COLUMN_ALIASES.items() for a in aliases
}


def _header_score(cells: Sequence[Any]) -> int:
    return len(
        {
            _ALIAS_TO_FIELD[k]
            for k in (_norm(_clean_header(c)) for c in cells)
            if k in _ALIAS_TO_FIELD
        }
    )


# --------------------------------------------------------------------------- #
# Excel reader / validator (pure logic below, easy to unit test)
# --------------------------------------------------------------------------- #
def read_excel_matrix(
    path: Path, sheet: Optional[Any] = None
) -> Tuple[List[str], List[List[str]], int]:
    """Read everything as text. Returns (headers, data rows, Excel row of the first data row)."""
    best: Optional[Tuple[int, str, int, pd.DataFrame]] = None
    matching: List[str] = []
    try:
        with pd.ExcelFile(path) as xl:
            sheets = [sheet] if sheet is not None else list(xl.sheet_names)
            for name in sheets:
                raw = xl.parse(
                    sheet_name=name,
                    header=None,
                    dtype=str,
                    keep_default_na=False,
                    na_filter=False,
                )
                if raw.empty:
                    continue
                top = min(HEADER_SCAN_ROWS, len(raw))
                scores = [_header_score(raw.iloc[i].tolist()) for i in range(top)]
                if not scores or max(scores) < MIN_HEADER_SCORE:
                    continue
                idx = scores.index(max(scores))
                matching.append(str(name))
                if best is None or scores[idx] > best[0]:
                    best = (scores[idx], str(name), idx, raw)
    except ImportError as exc:
        raise FileNotFoundError(
            f"Missing Excel engine ({exc}); for .xls add 'xlrd' (already installed)."
        ) from exc
    except Exception as exc:
        raise ValueError(f"Cannot read '{path.name}': {exc}") from exc

    if best is None:
        raise ValueError(
            "No sheet with a recognisable header row found (need at least "
            f"{MIN_HEADER_SCORE} of: {', '.join(v[0] for v in USER_COLUMN_ALIASES.values())} "
            f"within the first {HEADER_SCAN_ROWS} rows)."
        )
    _, sheet_name, header_idx, raw = best
    if len(matching) > 1:
        log.warning(
            "Several sheets look like user sheets (%s); using '%s' only.",
            ", ".join(matching),
            sheet_name,
        )
    headers = [_clean_header(v) for v in raw.iloc[header_idx].tolist()]
    body = raw.iloc[header_idx + 1 :].apply(lambda col: col.map(_clean_text))
    rows = body.values.tolist()
    log.info(
        "Sheet '%s': header at Excel row %d, %d data rows, %d columns",
        sheet_name,
        header_idx + 1,
        len(rows),
        len(headers),
    )
    return headers, rows, header_idx + 2


def classify_columns(
    headers: Sequence[str],
    rows: Sequence[Sequence[str]],
    issues: List[Issue],
    ignored: Sequence[str] = IGNORED_COLUMNS,
) -> ColumnMap:
    """Split headers into user columns, ';' authorization groups and dynamic roles."""
    ignored_norm = {_norm(i) for i in ignored}
    user_pos: Dict[str, int] = {}
    roles: Dict[str, List[int]] = {}
    groups: Dict[str, List[int]] = {}

    for pos, header in enumerate(headers):
        if header == "":
            if any(r[pos] != "" for r in rows):
                issues.append(
                    Issue(
                        "WARNING",
                        f"Column {get_column_letter(pos + 1)} has data but "
                        "no header - ignored.",
                    )
                )
            continue
        if header.startswith(";"):
            name = header[1:].strip()
            if not name:
                issues.append(
                    Issue(
                        "ERROR",
                        f"Column {get_column_letter(pos + 1)}: authorization "
                        "group header ';' has no name.",
                    )
                )
            else:
                groups.setdefault(name, []).append(pos)
            continue
        field_name = _ALIAS_TO_FIELD.get(_norm(header))
        if field_name:
            if field_name in user_pos:
                issues.append(
                    Issue(
                        "ERROR",
                        f"Column '{header}' appears twice "
                        f"({_col_letters([user_pos[field_name], pos])}).",
                    )
                )
            else:
                user_pos[field_name] = pos
            continue
        if _norm(header) in ignored_norm:
            if any(_mark(r[pos]) in ASSIGN_MARKS for r in rows):
                issues.append(
                    Issue(
                        "WARNING",
                        f"Column {get_column_letter(pos + 1)} '{header}' is on the "
                        "IGNORED_COLUMNS list but contains x marks - it is NOT sent "
                        "as a role. If it is a real role, remove it from IGNORED_COLUMNS.",
                    )
                )
            else:
                log.info(
                    "Column %s '%s' is not an access column - ignored.",
                    get_column_letter(pos + 1),
                    header,
                )
            continue
        roles.setdefault(header, []).append(pos)

    for label, mapping in (("Role", roles), ("Authorization group", groups)):
        for name, positions in mapping.items():
            if len(positions) > 1:
                issues.append(
                    Issue(
                        "WARNING",
                        f"{label} column '{name}' appears {len(positions)}x (columns "
                        f"{_col_letters(positions)}): treated as ONE {label.lower()}, "
                        "assigned if any of them is x.",
                    )
                )
    for field_name, aliases in USER_COLUMN_ALIASES.items():
        if field_name not in user_pos:
            issues.append(
                Issue(
                    "ERROR",
                    f"Required column missing: '{aliases[0]}'. "
                    f"Columns found: {list(headers)}",
                )
            )
    if not roles and not groups:
        issues.append(Issue("ERROR", "No role or authorization group columns found."))
    return ColumnMap(
        user_pos,
        {k: tuple(v) for k, v in roles.items()},
        {k: tuple(v) for k, v in groups.items()},
    )


def _marked(
    cells: Sequence[str],
    positions: Sequence[int],
    label: str,
    excel_row: int,
    issues: List[Issue],
) -> bool:
    assigned = False
    for pos in positions:
        mark = _mark(cells[pos])
        if mark in ASSIGN_MARKS:
            assigned = True
        elif mark not in IGNORE_MARKS:
            issues.append(
                Issue(
                    "ERROR",
                    f"Unexpected value '{cells[pos]}' in column "
                    f"{get_column_letter(pos + 1)} '{label}' "
                    "(use x, NA, N/A or blank).",
                    excel_row,
                )
            )
    return assigned


def parse_users(
    rows: Sequence[Sequence[str]], cmap: ColumnMap, first_row: int, issues: List[Issue]
) -> List[UserRecord]:
    """Group rows into users, validating each row and cross-row consistency."""
    users: Dict[str, UserRecord] = {}
    company_rows: Dict[Tuple[str, str], int] = {}
    ccgid_owner: Dict[str, Tuple[str, int]] = {}
    last_identity: Optional[Tuple[int, Dict[str, str]]] = None

    for offset, cells in enumerate(rows):
        excel_row = first_row + offset
        if not any(cells):
            continue
        vals = {f: cells[p] for f, p in cmap.user_pos.items()}
        vals["external_id"] = _clean_id(vals["external_id"])
        vals["company"] = _clean_id(vals["company"])
        if not any(vals.values()):
            issues.append(
                Issue(
                    "ERROR",
                    "Row has role/group values but no user columns "
                    "(CCGID, Username, ..., Company Code).",
                    excel_row,
                )
            )
            continue
        if (
            last_identity
            and vals["company"]
            and not any(vals[f] for f in IDENTITY_FIELDS)
        ):
            vals.update(last_identity[1])
            issues.append(
                Issue(
                    "WARNING",
                    "User columns are empty (merged cells?) - taken from "
                    f"row {last_identity[0]} ('{vals['user_name']}').",
                    excel_row,
                )
            )
        row_ok = True
        for f in USER_COLUMN_ALIASES:
            if not vals[f]:
                issues.append(Issue("ERROR", f"Missing {FIELD_LABELS[f]}.", excel_row))
                row_ok = False
        if vals["email"] and not EMAIL_RE.match(vals["email"]):
            issues.append(
                Issue("ERROR", f"Invalid email '{vals['email']}'.", excel_row)
            )

        access = AccessRow(excel_row, vals["company"])
        for role, positions in cmap.roles.items():
            if _marked(cells, positions, role, excel_row, issues):
                access.roles.append(role)
            if any(_mark(cells[pos]) in REMOVAL_MARKS for pos in positions):
                access.roles_to_remove.append(role)
        for group, positions in cmap.groups.items():
            if _marked(cells, positions, ";" + group, excel_row, issues):
                access.groups.append(group)
            if any(_mark(cells[pos]) in REMOVAL_MARKS for pos in positions):
                access.groups_to_remove.append(group)
        if not row_ok:
            continue
        last_identity = (excel_row, {f: vals[f] for f in IDENTITY_FIELDS})

        key = vals["user_name"].casefold()
        user = users.get(key)
        if user is None:
            user = users[key] = UserRecord(
                vals["user_name"],
                vals["external_id"],
                vals["first_name"],
                vals["last_name"],
                vals["email"],
                excel_row,
            )
        else:
            for f, mine, theirs in (
                ("external_id", user.external_id, vals["external_id"]),
                ("first_name", user.first_name, vals["first_name"]),
                ("last_name", user.last_name, vals["last_name"]),
                ("email", user.email.casefold(), vals["email"].casefold()),
            ):
                if mine != theirs:
                    issues.append(
                        Issue(
                            "ERROR",
                            f"User '{vals['user_name']}': {FIELD_LABELS[f]} "
                            f"differs from row {user.first_row}.",
                            excel_row,
                        )
                    )
        pair = (key, vals["company"].casefold())
        if pair in company_rows:
            issues.append(
                Issue(
                    "ERROR",
                    f"Duplicate user + Company Code ('{vals['user_name']}', "
                    f"'{vals['company']}') - same as row {company_rows[pair]}.",
                    excel_row,
                )
            )
        else:
            company_rows[pair] = excel_row
            user.rows.append(access)

        owner = ccgid_owner.setdefault(vals["external_id"].casefold(), (key, excel_row))
        if owner[0] != key:
            issues.append(
                Issue(
                    "ERROR",
                    f"CCGID '{vals['external_id']}' is already used by a "
                    f"different user (row {owner[1]}).",
                    excel_row,
                )
            )
        if (
            not access.roles
            and not access.groups
            and not access.roles_to_remove
            and not access.groups_to_remove
        ):
            issues.append(
                Issue(
                    "WARNING",
                    f"'{vals['user_name']}' / {vals['company']}: "
                    "no role or authorization group marked.",
                    excel_row,
                )
            )
    return list(users.values())


def is_allowed_email(email: str) -> bool:
    """True only for <local>@cma-cgm.com (exact domain, case-insensitive)."""
    local, sep, domain = email.strip().rpartition("@")
    return bool(local) and sep == "@" and domain.casefold() == ALLOWED_EMAIL_DOMAIN


def split_by_email_domain(
    users: Sequence[UserRecord],
) -> Tuple[List[UserRecord], List[RejectedUser]]:
    """(users with an allowed email, users rejected because of their email domain)."""
    valid: List[UserRecord] = []
    rejected: List[RejectedUser] = []
    for u in users:
        if is_allowed_email(u.email):
            valid.append(u)
        else:
            rows = tuple(sorted({u.first_row, *(r.excel_row for r in u.rows)}))
            rejected.append(RejectedUser(u.user_name, u.external_id, u.email, rows))
    return valid, rejected


def load_and_validate(
    path: Path, sheet: Optional[Any] = None
) -> Tuple[List[UserRecord], List[RejectedUser]]:
    """Read + validate. Raises ValueError listing the errors; nothing is sent before this passes.

    Returns (users to process, users rejected for a non-cma-cgm.com email).
    Rejected users are not fatal: they are reported and kept out of the payloads.
    """
    headers, rows, first_row = read_excel_matrix(path, sheet)
    issues: List[Issue] = []
    cmap = classify_columns(headers, rows, issues)
    users: List[UserRecord] = []
    if not any(i.severity == "ERROR" for i in issues):
        users = parse_users(rows, cmap, first_row, issues)
        if not users:
            issues.append(Issue("ERROR", "No user rows found below the header."))

    errors = [i for i in issues if i.severity == "ERROR"]
    for w in [i for i in issues if i.severity == "WARNING"][:200]:
        log.warning("%s", w)
    if errors:
        shown = "\n".join(str(e) for e in errors[:50])
        more = f"\n... {len(errors) - 50} more error(s)" if len(errors) > 50 else ""
        raise ValueError(
            f"{len(errors)} critical validation error(s) - nothing was sent to Medius.\n"
            f"{shown}{more}"
        )
    users, rejected = split_by_email_domain(users)
    for r in rejected:
        log.error("%s", r)
    if not users:
        shown = "\n".join(str(r) for r in rejected[:50])
        raise ValueError(
            f"No user with an @{ALLOWED_EMAIL_DOMAIN} email - nothing was sent to Medius.\n"
            f"{shown}"
        )
    log.info(
        "Validation passed: %d users, %d access rows, %d role columns, %d group columns"
        " (%d user(s) rejected for non-%s email)",
        len(users),
        sum(len(u.rows) for u in users),
        len(cmap.roles),
        len(cmap.groups),
        len(rejected),
        ALLOWED_EMAIL_DOMAIN,
    )
    return users, rejected


# --------------------------------------------------------------------------- #
# Payload builders (pure)
# --------------------------------------------------------------------------- #
def build_roles_payload(
    users: Sequence[UserRecord], env: EnvConfig, erp_source_id: str
) -> List[Dict[str, Any]]:
    """One entry per user (roles of all Company Code rows combined). Users with no role are left out."""
    payload: List[Dict[str, Any]] = []
    for u in users:
        roles = [
            {"roleCompany": row.company, "roleName": role}
            for row in u.rows
            for role in row.roles
        ]
        roles.extend(
            {
                "roleCompany": row.company,
                "roleName": role,
                "toBeRemoved": True,
            }
            for row in u.rows
            for role in row.roles_to_remove
        )
        if not roles:
            continue
        payload.append(
            {
                "erpSourceId": erp_source_id,
                "externalSystemId": u.external_id,
                "isActive": True,
                "userName": u.user_name,
                "firstName": u.first_name,
                "lastName": u.last_name,
                "email": u.email,
                "timezone": TIMEZONE,
                "language": LANGUAGE,
                "format": FORMAT,
                "roles": roles,
                "identityProviders": [
                    {
                        "identityProviderName": env.identity_provider,
                        "nameIdentifier": u.external_id,
                    }
                ],
            }
        )
    return payload


def build_auth_groups_payload(
    users: Sequence[UserRecord], erp_source_id: str
) -> List[Dict[str, Any]]:
    """One entry per user + Company Code row with at least one authorization group."""
    payload: List[Dict[str, Any]] = []
    for u in users:
        for row in u.rows:
            if not row.groups and not row.groups_to_remove:
                continue
            payload.append(
                {
                    "erpSourceId": erp_source_id,
                    "userName": u.user_name,
                    "companyId": row.company,
                    "authorizationGroups": [
                        {
                            "companyId": row.company,
                            "authorizationGroupName": g,
                            "toBeRemoved": False,
                        }
                        for g in row.groups
                    ]
                    + [
                        {
                            "companyId": row.company,
                            "authorizationGroupName": g,
                            "toBeRemoved": True,
                        }
                        for g in row.groups_to_remove
                    ],
                }
            )
    return payload


def to_json_bytes(payload: Any) -> bytes:
    return json.dumps(payload, ensure_ascii=False).encode("utf-8")


# --------------------------------------------------------------------------- #
# Settings from .env (never hard-coded, never logged)
# --------------------------------------------------------------------------- #
def load_env_config(name: str, dry_run: bool) -> EnvConfig:
    """Read MEDIUS_<SETTING>_<ENV> variables. A dry run tolerates missing values."""
    key = name.upper()
    settings = (
        "TOKEN_URL",
        "ROLES_URL",
        "AUTH_GROUPS_URL",
        "IDENTITY_PROVIDER",
        "CLIENT_ID",
        "CLIENT_SECRET",
    )
    values = {s: os.environ.get(f"MEDIUS_{s}_{key}", "").strip() for s in settings}
    missing = [f"MEDIUS_{s}_{key}" for s, v in values.items() if not v]
    if missing and not dry_run:
        raise ValueError(f"Missing .env variable(s): {', '.join(missing)}")
    if missing:
        log.warning("Dry run: .env variable(s) not set: %s", ", ".join(missing))
    return EnvConfig(
        name=key,
        token_url=values["TOKEN_URL"],
        roles_url=values["ROLES_URL"],
        auth_groups_url=values["AUTH_GROUPS_URL"],
        identity_provider=values["IDENTITY_PROVIDER"] or "<identity provider not set>",
        client_id=values["CLIENT_ID"],
        client_secret=values["CLIENT_SECRET"],
    )


# --------------------------------------------------------------------------- #
# HTTP client
# --------------------------------------------------------------------------- #
HTTP_HINTS: Mapping[int, str] = {
    400: "Bad request - payload rejected by Medius",
    401: "Unauthorized - token rejected",
    403: "Forbidden - client has no permission/scope for this endpoint",
    404: "Not found - check tenant / endpoint URL",
    406: "Not acceptable",
    409: "Conflict - request conflicts with current state",
    415: "Unsupported media type",
    429: "Too many requests (retries exhausted)",
    500: "Internal server error",
    502: "Bad gateway (retries exhausted)",
    503: "Service unavailable (retries exhausted)",
    504: "Gateway timeout (retries exhausted)",
}
_ERROR_KEYS = (
    "message",
    "Message",
    "error_description",
    "error",
    "detail",
    "title",
    "errors",
    "Errors",
)


def _safe_error_text(resp: requests.Response, limit: int = 400) -> str:
    try:
        body = resp.json()
    except ValueError:
        return (resp.text or "").strip().replace("\n", " ")[:limit]
    if isinstance(body, dict):
        found = {k: body[k] for k in _ERROR_KEYS if body.get(k)}
        if found:
            return json.dumps(found, ensure_ascii=False)[:limit]
    return json.dumps(body, ensure_ascii=False)[:limit]


def _find_import_link(resp: requests.Response) -> str:
    loc = resp.headers.get("Location", "")
    if loc:
        return loc
    text = (resp.text or "").strip().strip('"')
    if text.startswith("http"):
        return text
    try:
        body = resp.json()
    except ValueError:
        return ""
    if isinstance(body, dict):
        for value in body.values():
            if isinstance(value, str) and value.startswith("http"):
                return value
    return ""


def process_response(resp: requests.Response) -> ApiResult:
    """2xx = accepted into the Medius import queue (processing is asynchronous)."""
    code = resp.status_code
    if 200 <= code < 300:
        return ApiResult(
            True, code, "Accepted by Medius import queue", _find_import_link(resp)
        )
    hint = HTTP_HINTS.get(code, f"Unexpected HTTP status {code}")
    return ApiResult(False, code, f"{hint}. {_safe_error_text(resp)}".strip())


def _sleep_time(attempt: int, retry_after: Optional[str] = None) -> float:
    if retry_after and retry_after.strip().isdigit():
        return min(float(retry_after), BACKOFF_MAX)
    return min(BACKOFF_MAX, BACKOFF_BASE * (2**attempt)) + random.uniform(0, 0.5)


TLS_HELP = (
    "The HTTPS certificate could not be verified. On a company network this is usually "
    "a proxy re-signing traffic with a corporate root CA that Python does not trust. "
    "Fix: run 'pip install truststore' (uses the Windows certificate store), or set "
    "MEDIUS_CA_BUNDLE in .env to the path of your company CA .pem file."
)


def _tls_reason(exc: Exception) -> str:
    """Short SSL reason (e.g. CERTIFICATE_VERIFY_FAILED ...) without URLs or secrets."""
    match = re.search(r"\[SSL[^\]]*\][^'\")]*", str(exc))
    if not match:
        return type(exc).__name__
    return re.sub(r"\s*\(_ssl\.c:\d+", "", match.group(0)).strip()


def _configure_tls(session: requests.Session) -> None:
    """Optional corporate-CA support. Certificate checking is never switched off."""
    bundle = os.environ.get("MEDIUS_CA_BUNDLE", "").strip()
    if bundle:
        if not Path(bundle).is_file():
            raise ValueError(f"MEDIUS_CA_BUNDLE file not found: {bundle}")
        session.verify = bundle
        log.info("TLS: using CA bundle from MEDIUS_CA_BUNDLE")
        return
    try:
        import ssl

        import truststore
    except ImportError:
        log.warning(
            "TLS: package 'truststore' is not installed in this Python environment "
            "- using the built-in certificate list (pip install truststore)."
        )
        return

    ctx = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)

    class _OsTrustAdapter(HTTPAdapter):
        """Verify with the OS certificate store. Mounted on the session itself, so it
        works whatever order requests/truststore were imported in."""

        def init_poolmanager(self, *args: Any, **kwargs: Any) -> None:
            kwargs["ssl_context"] = ctx
            super().init_poolmanager(*args, **kwargs)

        def proxy_manager_for(self, *args: Any, **kwargs: Any) -> Any:
            kwargs["ssl_context"] = ctx
            return super().proxy_manager_for(*args, **kwargs)

    session.mount("https://", _OsTrustAdapter())
    log.info("TLS: using the operating-system certificate store (truststore)")


class MediusClient:
    """requests.Session wrapper: OAuth2 client credentials, retries with backoff, batching."""

    def __init__(self, env: EnvConfig) -> None:
        self.env = env
        self.session = requests.Session()
        _configure_tls(self.session)
        self._token: Optional[str] = None
        self._expires_at = 0.0

    # --- authentication ---
    def _fetch_token(self) -> str:
        form = {
            "grant_type": "client_credentials",
            "client_id": self.env.client_id,
            "client_secret": self.env.client_secret,
        }
        last = "no response"
        for attempt in range(MAX_RETRIES + 1):
            try:
                resp = self.session.post(
                    self.env.token_url,
                    data=form,
                    headers={"Accept": "application/json"},
                    timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
                )
            except requests.exceptions.SSLError as exc:
                last = f"TLS/SSL error ({_tls_reason(exc)}). {TLS_HELP}"
                break  # a certificate problem will not fix itself - do not retry
            except (requests.ConnectionError, requests.Timeout) as exc:
                last = f"network error ({type(exc).__name__})"
            else:
                if resp.status_code == 200:
                    try:
                        body = resp.json()
                        token = str(body["access_token"])
                    except (ValueError, KeyError, TypeError) as exc:
                        raise RuntimeError(
                            "Token response had no access_token."
                        ) from exc
                    ttl = float(body.get("expires_in", 3600) or 3600)
                    self._token = token
                    self._expires_at = time.monotonic() + max(ttl - TOKEN_SKEW, 30.0)
                    return token
                last = f"HTTP {resp.status_code}: {_safe_error_text(resp)}"
                if resp.status_code not in RETRY_STATUSES:
                    break
            if attempt < MAX_RETRIES:
                time.sleep(_sleep_time(attempt))
        raise RuntimeError(f"Medius authentication failed - {self._scrub(last)}")

    def _token_value(self) -> str:
        if self._token and time.monotonic() < self._expires_at:
            return self._token
        return self._fetch_token()

    def authenticate(self) -> None:
        self._token_value()

    def _scrub(self, text: str) -> str:
        for secret in (self.env.client_id, self.env.client_secret, self._token or ""):
            if secret:
                text = text.replace(secret, "***")
        return text

    # --- requests ---
    def post_json(self, url: str, payload: Any) -> ApiResult:
        body = to_json_bytes(payload)
        refreshed = False
        attempt = 0
        while True:
            headers = {
                "Authorization": f"Bearer {self._token_value()}",
                "Content-Type": "application/json; charset=utf-8",
                "Accept": "application/json",
            }
            try:
                resp = self.session.post(
                    url,
                    data=body,
                    headers=headers,
                    timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
                )
            except requests.exceptions.SSLError as exc:
                return ApiResult(
                    False, None, f"TLS/SSL error ({_tls_reason(exc)}). {TLS_HELP}"
                )
            except (requests.ConnectionError, requests.Timeout) as exc:
                if attempt < MAX_RETRIES:
                    delay = _sleep_time(attempt)
                    log.warning(
                        "Network error (%s); retry %d/%d in %.1fs",
                        type(exc).__name__,
                        attempt + 1,
                        MAX_RETRIES,
                        delay,
                    )
                    time.sleep(delay)
                    attempt += 1
                    continue
                return ApiResult(False, None, f"Network error: {type(exc).__name__}")

            if resp.status_code == 401 and not refreshed:
                refreshed = True
                self._token, self._expires_at = None, 0.0
                log.info("HTTP 401 - refreshing token and retrying once")
                continue
            if resp.status_code in RETRY_STATUSES and attempt < MAX_RETRIES:
                delay = _sleep_time(attempt, resp.headers.get("Retry-After"))
                log.warning(
                    "HTTP %d; retry %d/%d in %.1fs",
                    resp.status_code,
                    attempt + 1,
                    MAX_RETRIES,
                    delay,
                )
                time.sleep(delay)
                attempt += 1
                continue
            res = process_response(resp)
            return ApiResult(
                res.ok, res.status_code, self._scrub(res.message), res.import_link
            )

    def submit_batches(
        self, url: str, items: Sequence[Dict[str, Any]], batch_size: int, label: str
    ) -> List[ApiResult]:
        """POST items in batches; one ApiResult per item (same order).

        HTTP 400 on a multi-entry batch is bisected to isolate bad entries;
        401/403/404 are systemic, so the remaining batches are not sent.
        """
        results: List[Optional[ApiResult]] = [None] * len(items)
        systemic: List[ApiResult] = []

        def send(lo: int, hi: int) -> None:
            if systemic:
                blocked = ApiResult(
                    False,
                    systemic[0].status_code,
                    f"Not sent - aborted after HTTP {systemic[0].status_code}",
                )
                results[lo:hi] = [blocked] * (hi - lo)
                return
            res = self.post_json(url, list(items[lo:hi]))
            if not res.ok and res.status_code == 400 and hi - lo > 1:
                log.warning(
                    "%s: batch of %d rejected (400) - isolating failing entries",
                    label,
                    hi - lo,
                )
                mid = (lo + hi) // 2
                send(lo, mid)
                send(mid, hi)
                return
            if not res.ok and res.status_code in (401, 403, 404):
                systemic.append(res)
            results[lo:hi] = [res] * (hi - lo)

        for start in range(0, len(items), batch_size):
            end = min(start + batch_size, len(items))
            log.info(
                "%s: sending entries %d-%d of %d", label, start + 1, end, len(items)
            )
            send(start, end)
        return [r for r in results if r is not None]


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def _fmt_failure(res: ApiResult) -> str:
    return f"HTTP {res.status_code} - {res.message}" if res.status_code else res.message


def _apply_roles(
    items: Sequence[Dict[str, Any]],
    api: Sequence[ApiResult],
    results: Mapping[str, UserResult],
) -> None:
    for item, res in zip(items, api):
        ur = results[item["userName"]]
        ur.roles_status = S_OK if res.ok else S_FAIL
        ur.roles_http = res.status_code
        ur.roles_message = "" if res.ok else _fmt_failure(res)
        ur.roles_link = res.import_link if res.ok else ""


def _apply_groups(
    items: Sequence[Dict[str, Any]],
    api: Sequence[ApiResult],
    results: Mapping[str, UserResult],
) -> None:
    """A user can have several entries (one per Company Code): FAILED if any entry failed."""
    failures: Dict[str, List[str]] = {}
    links: Dict[str, List[str]] = {}
    for item, res in zip(items, api):
        ur = results[item["userName"]]
        if res.ok:
            if ur.groups_status != S_FAIL:
                ur.groups_status = S_OK
                ur.groups_http = ur.groups_http or res.status_code
            if res.import_link:
                links.setdefault(ur.user_name, []).append(res.import_link)
        else:
            if ur.groups_status != S_FAIL:
                ur.groups_http = res.status_code
            ur.groups_status = S_FAIL
            failures.setdefault(ur.user_name, []).append(
                f"{item['companyId']}: {_fmt_failure(res)}"
            )
    for name, msgs in failures.items():
        results[name].groups_message = " | ".join(msgs)
    for name, urls in links.items():
        results[name].groups_link = " ".join(dict.fromkeys(urls))


def run_updates(
    users: Sequence[UserRecord],
    env: EnvConfig,
    batch_size: int,
    roles_payload: Sequence[Dict[str, Any]],
    groups_payload: Sequence[Dict[str, Any]],
    dry_run: bool,
) -> Tuple[List[UserResult], str]:
    results: Dict[str, UserResult] = {
        u.user_name: UserResult(u.user_name, u.external_id, env.name) for u in users
    }

    if dry_run:
        for p in roles_payload:
            results[p["userName"]].roles_status = S_DRY
        for p in groups_payload:
            results[p["userName"]].groups_status = S_DRY
        ts = now_utc()
        for r in results.values():
            r.timestamp = ts
        return list(results.values()), "SKIPPED (dry run)"

    client = MediusClient(env)
    client.authenticate()
    log.info("Authentication successful (token cached and reused)")

    if roles_payload:
        api = client.submit_batches(env.roles_url, roles_payload, batch_size, "Roles")
        _apply_roles(roles_payload, api, results)

    blocked = {
        p["userName"]
        for p in groups_payload
        if results[p["userName"]].roles_status == S_FAIL
    }
    for name in blocked:
        results[name].groups_status = S_SKIP
        results[name].groups_message = "Skipped because the Roles update failed"
    to_send = [p for p in groups_payload if p["userName"] not in blocked]
    if to_send:
        api = client.submit_batches(
            env.auth_groups_url, to_send, batch_size, "Authorization Groups"
        )
        _apply_groups(to_send, api, results)

    ts = now_utc()
    for r in results.values():
        r.timestamp = ts
    return list(results.values()), "SUCCESS"


# --------------------------------------------------------------------------- #
# Reporting
# --------------------------------------------------------------------------- #
def _reason(status: str, message: str, limit: int = 140) -> str:
    return f"{status} - {message[:limit]}" if status in (S_FAIL, S_SKIP) else status


def results_frame(results: Sequence[UserResult]) -> pd.DataFrame:
    rows = []
    for r in results:
        if r.roles_status == S_REJ:
            errors = r.roles_message
        else:
            errors = "; ".join(
                filter(
                    None,
                    (
                        f"Roles: {r.roles_message}"
                        if r.roles_status in (S_FAIL, S_SKIP)
                        else "",
                        f"Auth Groups: {r.groups_message}"
                        if r.groups_status in (S_FAIL, S_SKIP)
                        else "",
                    ),
                )
            )
        rows.append(
            {
                "Username": r.user_name,
                "External ID": r.external_id,
                "Environment": r.environment,
                "Roles Status": r.roles_status,
                "Authorization Group Status": r.groups_status,
                "Roles HTTP Status": r.roles_http if r.roles_http is not None else "",
                "Auth Groups HTTP Status": r.groups_http
                if r.groups_http is not None
                else "",
                "Error Message": errors,
                "Timestamp": r.timestamp,
                "Roles Import Link": r.roles_link,
                "Auth Groups Import Link": r.groups_link,
            }
        )
    return pd.DataFrame(rows)


def build_summary(
    env: EnvConfig,
    users: Sequence[UserRecord],
    results: Sequence[UserResult],
    auth_status: str,
    dry_run: bool,
    rejected: Sequence[RejectedUser] = (),
) -> str:
    def count(attr: str, status: str) -> int:
        return sum(1 for r in results if getattr(r, attr) == status)

    lines = [
        f"Environment: {env.name}"
        + ("   (DRY RUN - no API calls made)" if dry_run else ""),
        f"Users found: {len(users)}   (access rows: {sum(len(u.rows) for u in users)})",
        f"Authentication: {auth_status}",
    ]
    for title, attr in (
        ("Roles", "roles_status"),
        ("Authorization Groups", "groups_status"),
    ):
        if dry_run:
            lines.append(
                f"{title}: prepared {count(attr, S_DRY)}, "
                f"not applicable {count(attr, S_NA)}"
            )
        else:
            lines.append(
                f"{title}: success {count(attr, S_OK)}, failed {count(attr, S_FAIL)}, "
                f"skipped {count(attr, S_SKIP)}, not applicable {count(attr, S_NA)}"
            )
    if rejected:
        lines.append(
            f"ERROR: {len(rejected)} user(s) rejected - email is not @{ALLOWED_EMAIL_DOMAIN} "
            "(excluded from payloads):"
        )
        for rj in list(rejected)[:MAX_SUMMARY_LINES]:
            rows = ", ".join(str(x) for x in rj.excel_rows)
            lines.append(f"  Row {rows}: {rj.user_name} <{rj.email}>")
        if len(rejected) > MAX_SUMMARY_LINES:
            lines.append(
                f"  ... {len(rejected) - MAX_SUMMARY_LINES} more - see the Excel report"
            )
    lines.append("")
    width = max((len(r.user_name) for r in results), default=10)
    for r in list(results)[:MAX_SUMMARY_LINES]:
        lines.append(
            f"{r.user_name:<{width}} | Roles: {_reason(r.roles_status, r.roles_message)}"
            f" | Auth Groups: {_reason(r.groups_status, r.groups_message)}"
        )
    if len(results) > MAX_SUMMARY_LINES:
        lines.append(
            f"... {len(results) - MAX_SUMMARY_LINES} more user(s) - see the Excel report"
        )
    failed = sum(1 for r in results if S_FAIL in (r.roles_status, r.groups_status))
    lines.append(
        "RESULT: "
        + ("SUCCESS" if not failed else f"COMPLETED WITH {failed} FAILED USER(S)")
        + (f"; {len(rejected)} USER(S) REJECTED (EMAIL DOMAIN)" if rejected else "")
    )
    return "\n".join(lines)


def save_payload_previews(
    roles: Sequence[Dict[str, Any]],
    groups: Sequence[Dict[str, Any]],
    env: EnvConfig,
    out_dir: Path,
    stamp: str,
) -> List[Path]:
    paths: List[Path] = []
    for label, payload in (("roles", roles), ("auth_groups", groups)):
        path = out_dir / f"Medius_{label}_payload_{env.name}_{stamp}.json"
        path.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        paths.append(path)
        log.info(
            "Dry run: %s payload with %d entr(y/ies) written to %s",
            label,
            len(payload),
            path.name,
        )
    return paths


# --------------------------------------------------------------------------- #
# Job entry points
# --------------------------------------------------------------------------- #
def add_arguments(p) -> None:
    p.add_argument(
        "--file", default=None, help="Excel access matrix (.xlsx/.xlsm/.xls)"
    )
    p.add_argument(
        "--pick",
        action="store_true",
        help="open a file-picker window to choose the Excel file (instead of --file)",
    )
    p.add_argument(
        "--env",
        required=True,
        choices=["QA", "PROD", "qa", "prod"],
        help="Medius environment",
    )
    p.add_argument(
        "--sheet", default=None, help="sheet name or number (default: auto-detect)"
    )
    p.add_argument(
        "--erp-source-id", default="IT_CAP", help="erpSourceId in the payloads"
    )
    p.add_argument("--batch-size", type=int, default=50)
    p.add_argument(
        "--confirm-production",
        action="store_true",
        help="required together with --execute when --env PROD",
    )


def resolve_input_file(file: Optional[str], pick: bool) -> Path:
    """--file path, or the Windows file-picker with --pick."""
    if file and pick:
        raise ValueError("Use either --file or --pick, not both.")
    if file:
        path = Path(file)
    elif pick:
        chosen = pick_files()
        if not chosen:
            raise ValueError("No file selected.")
        if len(chosen) > 1:
            log.warning("Several files selected; using the first: %s", chosen[0].name)
        path = Path(chosen[0])
    else:
        raise ValueError(
            "Give the Excel file with --file <path> or choose it with --pick."
        )
    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")
    if path.suffix.lower() not in ALLOWED_SUFFIXES:
        raise ValueError(
            f"Unsupported file type '{path.suffix}' (expected .xlsx/.xlsm/.xls)"
        )
    return path


def run(args) -> JobResult:
    dry_run = not args.execute
    path = resolve_input_file(args.file, args.pick)
    if args.batch_size < 1:
        raise ValueError("--batch-size must be at least 1")
    sheet: Optional[Any] = args.sheet
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    env = load_env_config(args.env, dry_run)
    if not dry_run and env.name == "PROD" and not args.confirm_production:
        raise ValueError(
            "PROD with --execute also needs --confirm-production. Nothing was changed."
        )

    users, rejected = load_and_validate(path, sheet)
    roles_payload = build_roles_payload(users, env, args.erp_source_id)
    groups_payload = build_auth_groups_payload(users, args.erp_source_id)
    log.info(
        "Payloads built: %d role entr(y/ies), %d authorization group entr(y/ies)",
        len(roles_payload),
        len(groups_payload),
    )

    out_dir = get_settings().output_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    extra: List[Path] = (
        save_payload_previews(roles_payload, groups_payload, env, out_dir, stamp)
        if dry_run
        else []
    )

    results, auth_status = run_updates(
        users, env, args.batch_size, roles_payload, groups_payload, dry_run
    )
    ts = now_utc()
    rejected_results = [
        UserResult(
            rj.user_name,
            rj.external_id,
            env.name,
            roles_status=S_REJ,
            roles_message=rj.reason,
            groups_status=S_REJ,
            groups_message=rj.reason,
            timestamp=ts,
        )
        for rj in rejected
    ]
    report = write_excel(
        {"Results": results_frame([*results, *rejected_results])},
        out_dir / f"Medius_User_Update_Report_{env.name}_{stamp}.xlsx",
    )
    return JobResult(
        build_summary(env, users, results, auth_status, dry_run, rejected),
        [report, *extra],
    )
