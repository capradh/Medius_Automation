import pytest

from automation.jobs.medius_user_access import (
    EnvConfig,
    build_auth_groups_payload,
    build_roles_payload,
    classify_columns,
    parse_users,
)

HEADERS = [
    "CCGID",
    "Username",
    "First Name",
    "Last Name",
    "Email",
    "Company Code",
    "Buyer",
    "Approver",
    ";GrpA",
    "Comments",
]
ENV = EnvConfig("QA", "t", "r", "g", "idp-1")


def _parse(rows):
    issues = []
    cmap = classify_columns(HEADERS, rows, issues)
    users = parse_users(rows, cmap, 2, issues)
    return users, issues


def _errors(issues):
    return [i for i in issues if i.severity == "ERROR"]


def test_basic_user_and_float_id():
    rows = [
        [
            "992025.0",
            "jdoe",
            "John",
            "Doe",
            "a@company.com",
            "0089",
            "x",
            "N/A",
            "X",
            "hi",
        ]
    ]
    users, issues = _parse(rows)
    assert not _errors(issues)
    assert users[0].external_id == "992025"
    assert users[0].rows[0].roles == ["Buyer"]
    assert users[0].rows[0].groups == ["GrpA"]


def test_roles_payload_combines_company_rows():
    rows = [
        ["1", "jdoe", "John", "Doe", "a@company.com", "0089", "x", "", "", ""],
        ["1", "jdoe", "John", "Doe", "a@company.com", "0090", "", "x", "x", ""],
    ]
    users, issues = _parse(rows)
    assert not _errors(issues)
    roles = build_roles_payload(users, ENV, "IT_CAP")
    assert len(roles) == 1
    assert roles[0]["roles"] == [
        {"roleCompany": "0089", "roleName": "Buyer"},
        {"roleCompany": "0090", "roleName": "Approver"},
    ]
    assert roles[0]["identityProviders"][0]["nameIdentifier"] == "1"
    groups = build_auth_groups_payload(users, "IT_CAP")
    assert [g["companyId"] for g in groups] == ["0090"]


def test_user_without_roles_not_in_roles_payload():
    rows = [["1", "jdoe", "John", "Doe", "a@company.com", "0089", "", "", "x", ""]]
    users, _ = _parse(rows)
    assert build_roles_payload(users, ENV, "IT_CAP") == []
    assert len(build_auth_groups_payload(users, "IT_CAP")) == 1


def test_fullwidth_x_assigns():
    rows = [["1", "jdoe", "John", "Doe", "a@company.com", "0089", "\uff38", "", "", ""]]
    users, issues = _parse(rows)
    assert not _errors(issues)
    assert users[0].rows[0].roles == ["Buyer"]


def test_unexpected_mark_is_error():
    rows = [["1", "jdoe", "John", "Doe", "a@company.com", "0089", "yes", "", "", ""]]
    _, issues = _parse(rows)
    assert any("Unexpected value" in i.message for i in _errors(issues))


def test_duplicate_user_company_is_error():
    row = ["1", "jdoe", "John", "Doe", "a@company.com", "0089", "x", "", "", ""]
    _, issues = _parse([row, list(row)])
    assert any("Duplicate user" in i.message for i in _errors(issues))


def test_invalid_email_is_error():
    rows = [["1", "jdoe", "John", "Doe", "not-an-email", "0089", "x", "", "", ""]]
    _, issues = _parse(rows)
    assert any("Invalid email" in i.message for i in _errors(issues))


def test_same_ccgid_different_user_is_error():
    rows = [
        ["1", "jdoe", "John", "Doe", "a@company.com", "0089", "x", "", "", ""],
        ["1", "asmith", "Ann", "Smith", "b@company.com", "0089", "x", "", "", ""],
    ]
    _, issues = _parse(rows)
    assert any("already used" in i.message for i in _errors(issues))


def test_merged_user_cells_reuse_previous_user():
    rows = [
        ["1", "jdoe", "John", "Doe", "a@company.com", "0089", "x", "", "", ""],
        ["", "", "", "", "", "0090", "x", "", "", ""],
    ]
    users, issues = _parse(rows)
    assert not _errors(issues)
    assert len(users) == 1 and len(users[0].rows) == 2


def test_missing_required_column_names_found_columns():
    headers = ["CCGID", "Username", "Buyer"]
    issues = []
    classify_columns(headers, [], issues)
    assert any(
        "Required column missing" in i.message and "Buyer" in i.message
        for i in issues
        if i.severity == "ERROR"
    )
