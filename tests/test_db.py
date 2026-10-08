from unittest.mock import MagicMock, patch

import pytest

ENV = {"DB_LARA_USER": "u", "DB_LARA_PASSWORD": "p", "DB_LARA_DSN": "h:1521/S"}


@pytest.fixture
def fake_conn(monkeypatch):
    for k, v in ENV.items():
        monkeypatch.setenv(k, v)
    conn = MagicMock()
    cur = conn.cursor.return_value.__enter__.return_value
    cur.description = [("ID",), ("NAME",)]
    cur.fetchall.return_value = [(1, "a"), (2, "b")]
    cur.rowcount = 2
    with patch("automation.db.oracledb.connect", return_value=conn):
        yield conn


def test_fetch_all_returns_dicts(fake_conn):
    from automation.db import fetch_all
    assert fetch_all("lara", "SELECT 1") == [{"ID": 1, "NAME": "a"}, {"ID": 2, "NAME": "b"}]
    fake_conn.close.assert_called_once()


def test_dry_run_rolls_back(fake_conn):
    from automation.db import execute_write
    assert execute_write("lara", "DELETE FROM t", dry_run=True) == 2
    fake_conn.rollback.assert_called_once()
    fake_conn.commit.assert_not_called()


def test_execute_commits(fake_conn):
    from automation.db import execute_write
    execute_write("lara", "DELETE FROM t", dry_run=False)
    fake_conn.commit.assert_called_once()
    fake_conn.rollback.assert_not_called()
