from unittest.mock import MagicMock, patch

import pytest

from automation.config import ConfigError, SmtpSettings, get_smtp_settings
from automation.notifications import build_message, send_email


def settings(**kw):
    base = dict(host="smtp.test", port=587, security="starttls", user="u", password="p",
                sender="me@test.com", default_to=("a@test.com",))
    base.update(kw)
    return SmtpSettings(**base)


def test_env_parsing():
    s = get_smtp_settings({"SMTP_HOST": "h", "SMTP_USER": "me@x.com", "EMAIL_TO": "a@x.com; b@x.com"})
    assert s.sender == "me@x.com" and s.default_to == ("a@x.com", "b@x.com") and s.port == 587
    with pytest.raises(ConfigError):
        get_smtp_settings({"SMTP_SECURITY": "weird"})


def test_attachment_built(tmp_path):
    f = tmp_path / "r.xlsx"
    f.write_bytes(b"data")
    msg = build_message("s", "b", "me@x.com", ["a@x.com"], attachments=[f])
    assert [p.get_filename() for p in msg.iter_attachments()] == ["r.xlsx"]


def test_send_starttls(tmp_path):
    server = MagicMock()
    with patch("automation.notifications.smtplib.SMTP") as smtp:
        smtp.return_value.__enter__.return_value = server
        send_email("subj", "body", settings=settings())
    server.starttls.assert_called_once()
    server.login.assert_called_once_with("u", "p")
    assert server.send_message.call_args.kwargs["to_addrs"] == ["a@test.com"]


def test_send_ssl_and_plain_relay():
    server = MagicMock()
    with patch("automation.notifications.smtplib.SMTP_SSL") as ssl_cls:
        ssl_cls.return_value.__enter__.return_value = server
        send_email("s", "b", to="x@test.com", settings=settings(security="ssl", port=465))
    ssl_cls.assert_called_once()
    server2 = MagicMock()
    with patch("automation.notifications.smtplib.SMTP") as smtp:
        smtp.return_value.__enter__.return_value = server2
        send_email("s", "b", settings=settings(security="none", port=25, user=None, password=None))
    server2.starttls.assert_not_called()
    server2.login.assert_not_called()


def test_missing_config_errors():
    with pytest.raises(ConfigError, match="SMTP_HOST"):
        send_email("s", "b", settings=settings(host=None))
    with pytest.raises(ConfigError, match="recipients"):
        send_email("s", "b", settings=settings(default_to=()))
