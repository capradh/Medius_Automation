"""Email notifications over SMTP (STARTTLS, SSL or plain relay)."""
from __future__ import annotations

import logging
import mimetypes
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable

from .config import ConfigError, SmtpSettings, get_smtp_settings, split_addresses

log = logging.getLogger(__name__)
MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024


def build_message(subject: str, body: str, sender: str, to: list[str], cc: list[str] | None = None,
                  attachments: Iterable[str | Path] | None = None, html: bool = False) -> EmailMessage:
    msg = EmailMessage()
    msg["Subject"], msg["From"], msg["To"] = subject, sender, ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    if html:
        msg.set_content("This message requires an HTML-capable email client.")
        msg.add_alternative(body, subtype="html")
    else:
        msg.set_content(body)

    total = 0
    for att in attachments or []:
        p = Path(att)
        data = p.read_bytes()
        total += len(data)
        if total > MAX_ATTACHMENT_BYTES:
            raise ValueError("Attachments exceed 20 MB; send a link or fewer files instead.")
        ctype, _ = mimetypes.guess_type(p.name)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=p.name)
    return msg


def send_email(subject: str, body: str, to: list[str] | str | None = None, cc: list[str] | None = None,
               attachments: Iterable[str | Path] | None = None, html: bool = False,
               settings: SmtpSettings | None = None) -> None:
    s = settings or get_smtp_settings()
    recipients = split_addresses(to) if isinstance(to, str) else list(to or s.default_to)
    if not s.host:
        raise ConfigError("SMTP_HOST is not set in .env")
    if not s.sender:
        raise ConfigError("Set EMAIL_FROM (or SMTP_USER) in .env")
    if not recipients:
        raise ConfigError("No recipients: pass --email-to or set EMAIL_TO in .env")

    msg = build_message(subject, body, s.sender, recipients, cc, attachments, html)
    ctx = ssl.create_default_context()
    log.info("Sending email via %s:%s (%s) to %s", s.host, s.port, s.security, ", ".join(recipients))

    server = (smtplib.SMTP_SSL(s.host, s.port, timeout=s.timeout, context=ctx)
              if s.security == "ssl" else smtplib.SMTP(s.host, s.port, timeout=s.timeout))
    with server as conn:
        if s.security == "starttls":
            conn.ehlo()
            conn.starttls(context=ctx)
            conn.ehlo()
        if s.user and s.password:
            conn.login(s.user, s.password)
        conn.send_message(msg, to_addrs=recipients + (cc or []))
    log.info("Email sent")
