from automation.connectivity import hint_for, parse_host_port


def test_easy_connect():
    assert parse_host_port("db.example.com:1522/SVC") == ("db.example.com", 1522)
    assert parse_host_port("db.example.com/SVC") == ("db.example.com", 1521)
    assert parse_host_port("//db.example.com:1521/SVC") == ("db.example.com", 1521)


def test_tns_descriptor():
    dsn = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=h.example.com)(PORT=1530))(CONNECT_DATA=(SERVICE_NAME=X)))"
    assert parse_host_port(dsn) == ("h.example.com", 1530)


def test_tns_alias_returns_none():
    assert parse_host_port("MYALIAS") is None


def test_hints():
    assert "password" in hint_for(Exception("ORA-01017: invalid credential")).lower()
    assert "listener" in hint_for(Exception("ORA-12541: no listener")).lower()
