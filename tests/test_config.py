import pytest

from automation.config import ConfigError, discover_databases, get_database, list_incomplete

ENV = {
    "DB_CORSAIR_USER": "u1", "DB_CORSAIR_PASSWORD": "p1", "DB_CORSAIR_DSN": "h1:1521/S1",
    "DB_CORSAIR_ENV": "prod",
    "DB_LARA_USER": "u2", "DB_LARA_PASSWORD": "p2", "DB_LARA_DSN": "h2:1521/S2",
    "DB_MY_DB_USER": "u3", "DB_MY_DB_PASSWORD": "p3", "DB_MY_DB_DSN": "h3/S3",
    "DB_BROKEN_USER": "x",
}


def test_discovers_complete_databases_only():
    dbs = discover_databases(ENV)
    assert set(dbs) == {"CORSAIR", "LARA", "MY_DB"}


def test_env_flag_and_default():
    dbs = discover_databases(ENV)
    assert dbs["CORSAIR"].is_prod
    assert not dbs["LARA"].is_prod


def test_get_database_case_insensitive():
    assert get_database("corsair", ENV).dsn == "h1:1521/S1"


def test_password_not_in_repr():
    assert "p1" not in repr(get_database("corsair", ENV))


def test_incomplete_reported():
    assert list_incomplete(ENV) == {"BROKEN": ["PASSWORD", "DSN"]}
    with pytest.raises(ConfigError, match="DB_BROKEN_PASSWORD"):
        get_database("broken", ENV)


def test_unknown_lists_available():
    with pytest.raises(ConfigError, match="CORSAIR"):
        get_database("nope", ENV)
