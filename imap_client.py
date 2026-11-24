"""Lightweight stdlib IMAP client for reading recovery mailboxes (e.g. Firstmail) and extracting verification codes."""

from __future__ import annotations

import email
import html
import imaplib
import re
from datetime import timezone
from email.header import decode_header
from email.message import Message
from email.utils import parseaddr, parsedate_to_datetime
from typing import Any

OUTLOOK_IMAP_HOST = "outlook.office365.com"
OUTLOOK_IMAP_PORT = 993

KNOWN_IMAP_HOSTS: dict[str, str] = {
    "rambler.ru": "imap.rambler.ru",
    "lenta.ru": "imap.rambler.ru",
    "autorambler.ru": "imap.rambler.ru",
    "myrambler.ru": "imap.rambler.ru",
    "ro.ru": "imap.rambler.ru",
    "mail.ru": "imap.mail.ru",
    "inbox.ru": "imap.mail.ru",
    "list.ru": "imap.mail.ru",
    "bk.ru": "imap.mail.ru",
    "yandex.ru": "imap.yandex.ru",
    "ya.ru": "imap.yandex.ru",
    "gmx.com": "imap.gmx.com",
    "gmx.de": "imap.gmx.net",
}

DEFAULT_RECOVERY_IMAP_HOST = "imap.firstmail.ltd"

_STYLE_SCRIPT_RE = re.compile(
    r"<(script|style|head|title)[^>]*>.*?</\1>", re.IGNORECASE | re.DOTALL
)
_TAG_RE = re.compile(r"<[^>]+>")
_URL_RE = re.compile(r"https?://\S+", re.IGNORECASE)
_CODE_NEAR_KEYWORD_RE = re.compile(
    r"(?:security\s*code|verification\s*code|single[- ]use\s*code|one[- ]time\s*code|"
    r"код\s*(?:безопасности|подтверждения)|code|код)[^0-9A-Za-z]{0,25}(\d{4,8})\b",
    re.IGNORECASE,
)
_STANDALONE_CODE_RE = re.compile(r"(?<!\d)(\d{6,7})(?!\d)")


class RecoveryMailError(Exception):
    """Raised when connecting to or reading the recovery IMAP mailbox fails."""


def resolve_imap_host(email_addr: str, override_host: str = "") -> str:
    if override_host.strip():
        return override_host.strip()
    domain = email_addr.strip().lower().split("@")[-1] if "@" in email_addr else ""
    return KNOWN_IMAP_HOSTS.get(domain, DEFAULT_RECOVERY_IMAP_HOST)


def _decode_mime_header(value: str | None) -> str:
    if not value:
        return ""
    parts: list[str] = []
    for chunk, charset in decode_header(value):
        if isinstance(chunk, bytes):
            enc = charset or "utf-8"
            try:
                parts.append(chunk.decode(enc, errors="replace"))
            except LookupError:
                parts.append(chunk.decode("utf-8", errors="replace"))
        else:
            parts.append(str(chunk))
    return "".join(parts).strip()


def _html_to_text(raw_html: str) -> str:
    cleaned = _STYLE_SCRIPT_RE.sub(" ", raw_html)
    cleaned = re.sub(r"<br\s*/?>", "\n", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"</(p|div|tr|li|h[1-6])>", "\n", cleaned, flags=re.IGNORECASE)
    cleaned = _TAG_RE.sub(" ", cleaned)
    cleaned = html.unescape(cleaned)
    cleaned = re.sub(r"[ \t]+", " ", cleaned)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()


def _extract_body_text(msg: Message) -> str:
    plain_parts: list[str] = []
    html_parts: list[str] = []

    if msg.is_multipart():
        for part in msg.walk():
            disposition = str(part.get("Content-Disposition") or "").lower()
            if "attachment" in disposition:
                continue
            ctype = part.get_content_type()
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            charset = part.get_content_charset() or "utf-8"
            try:
                decoded = payload.decode(charset, errors="replace")
            except LookupError:
                decoded = payload.decode("utf-8", errors="replace")
            if ctype == "text/plain":
                plain_parts.append(decoded)
            elif ctype == "text/html":
                html_parts.append(decoded)
    else:
        payload = msg.get_payload(decode=True)
        if payload is not None:
            charset = msg.get_content_charset() or "utf-8"
            try:
                decoded = payload.decode(charset, errors="replace")
            except LookupError:
                decoded = payload.decode("utf-8", errors="replace")
            if msg.get_content_type() == "text/html":
                html_parts.append(decoded)
            else:
                plain_parts.append(decoded)

    if plain_parts:
        return "\n".join(p.strip() for p in plain_parts if p.strip()).strip()
    if html_parts:
        return _html_to_text("\n".join(html_parts))
    return ""


def extract_verification_codes(subject: str, body_text: str) -> list[str]:
    """Extract likely OTP / verification codes (4–8 digits) from email subject and text."""
    combined = f"{subject}\n{_URL_RE.sub(' ', body_text)}"
    found: list[str] = []

    for match in _CODE_NEAR_KEYWORD_RE.finditer(combined):
        code = match.group(1)
        if code not in found:
            found.append(code)

    if not found:
        for match in _STANDALONE_CODE_RE.finditer(combined):
            code = match.group(1)
            if code not in found:
                found.append(code)

    return found[:5]


def fetch_recovery_messages(
    *,
    email_addr: str,
    password: str,
    imap_host: str = "",
    imap_port: int = 993,
    limit: int = 8,
    timeout: int = 15,
) -> dict[str, Any]:
    """Connect over IMAP-SSL to the recovery mailbox and return latest messages with extracted codes."""
    host = resolve_imap_host(email_addr, imap_host)
    client: imaplib.IMAP4_SSL | None = None
    try:
        client = imaplib.IMAP4_SSL(host, imap_port, timeout=timeout)
        client.login(email_addr.strip(), password.strip())
    except imaplib.IMAP4.error as exc:
        err_text = str(exc)
        if "AUTH" in err_text.upper() or "LOGIN" in err_text.upper():
            raise RecoveryMailError(
                f"Неверный логин или пароль для {email_addr} на {host}"
            ) from exc
        raise RecoveryMailError(f"Ошибка IMAP ({host}): {err_text}") from exc
    except Exception as exc:
        raise RecoveryMailError(f"Не удалось подключиться к {host}:{imap_port} ({exc})") from exc

    messages: list[dict[str, Any]] = []
    try:
        for folder_name in ("INBOX", "Junk", "Spam"):
            status, _ = client.select(folder_name, readonly=True)
            if status != "OK":
                continue
            status, search_data = client.search(None, "ALL")
            if status != "OK" or not search_data or not search_data[0]:
                continue
            msg_nums = search_data[0].split()
            recent_nums = list(reversed(msg_nums[-limit:]))

            for num in recent_nums:
                f_status, raw_data = client.fetch(num, "(RFC822)")
                if f_status != "OK" or not raw_data:
                    continue
                raw_bytes = None
                for part in raw_data:
                    if isinstance(part, tuple) and len(part) >= 2 and isinstance(part[1], bytes):
                        raw_bytes = part[1]
                        break
                if not raw_bytes:
                    continue

                parsed = email.message_from_bytes(raw_bytes)
                subject = _decode_mime_header(parsed.get("Subject")) or "(Без темы)"
                sender = _decode_mime_header(parsed.get("From"))
                date_str = _decode_mime_header(parsed.get("Date"))
                body_text = _extract_body_text(parsed)
                preview = re.sub(r"\s+", " ", body_text)[:220].strip()
                codes = extract_verification_codes(subject, body_text)

                messages.append(
                    {
                        "folder": folder_name,
                        "subject": subject,
                        "sender": sender,
                        "date": date_str,
                        "preview": preview,
                        "text": body_text[:6000],
                        "codes": codes,
                    }
                )
            # If INBOX already yielded messages, no need to keep scanning empty folders unless we want junk too
            if folder_name == "INBOX" and len(messages) >= limit:
                break
    finally:
        try:
            client.logout()
        except Exception:
            pass

    return {
        "recovery_email": email_addr.strip().lower(),
        "imap_host": host,
        "messages": messages[:limit],
    }


def _extract_body_rich(msg: Message) -> tuple[str, str, str]:
    """Return (body_content, content_type ['html'|'text'], plain_preview)."""
    plain_parts: list[str] = []
    html_parts: list[str] = []

    if msg.is_multipart():
        for part in msg.walk():
            disposition = str(part.get("Content-Disposition") or "").lower()
            if "attachment" in disposition:
                continue
            ctype = part.get_content_type()
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            charset = part.get_content_charset() or "utf-8"
            try:
                decoded = payload.decode(charset, errors="replace")
            except LookupError:
                decoded = payload.decode("utf-8", errors="replace")
            if ctype == "text/html":
                html_parts.append(decoded)
            elif ctype == "text/plain":
                plain_parts.append(decoded)
    else:
        payload = msg.get_payload(decode=True)
        if payload is not None:
            charset = msg.get_content_charset() or "utf-8"
            try:
                decoded = payload.decode(charset, errors="replace")
            except LookupError:
                decoded = payload.decode("utf-8", errors="replace")
            if msg.get_content_type() == "text/html":
                html_parts.append(decoded)
            else:
                plain_parts.append(decoded)

    plain_text = "\n".join(p.strip() for p in plain_parts if p.strip()).strip()
    html_text = "\n".join(h.strip() for h in html_parts if h.strip()).strip()
    if html_text:
        preview_src = plain_text or _html_to_text(html_text)
        return html_text, "html", re.sub(r"\s+", " ", preview_src)[:220].strip()
    return plain_text, "text", re.sub(r"\s+", " ", plain_text)[:220].strip()


def _parse_iso_date(date_header: str) -> str:
    if not date_header:
        return ""
    try:
        dt = parsedate_to_datetime(date_header)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    except Exception:
        return ""


def fetch_outlook_imap_messages(
    *,
    email_addr: str,
    access_token: str,
    folder: str = "inbox",
    top: int = 30,
    timeout: int = 20,
) -> list[dict[str, Any]]:
    """Fetch messages from Outlook.com via IMAP XOAUTH2 (outlook.office365.com:993)."""
    imap_folders = ["INBOX"] if folder == "inbox" else ["Junk", "Junk Email", "Spam"]
    auth_string = f"user={email_addr.strip()}\x01auth=Bearer {access_token}\x01\x01".encode("utf-8")

    client: imaplib.IMAP4_SSL | None = None
    try:
        client = imaplib.IMAP4_SSL(OUTLOOK_IMAP_HOST, OUTLOOK_IMAP_PORT, timeout=timeout)
        client.authenticate("XOAUTH2", lambda _challenge: auth_string)
    except imaplib.IMAP4.error as exc:
        raise RecoveryMailError(f"IMAP XOAUTH2 отклонён: {exc}") from exc
    except Exception as exc:
        raise RecoveryMailError(f"Ошибка подключения к {OUTLOOK_IMAP_HOST}: {exc}") from exc

    items: list[dict[str, Any]] = []
    try:
        selected = False
        for fname in imap_folders:
            status, _ = client.select(f'"{fname}"', readonly=True)
            if status == "OK":
                selected = True
                break
        if not selected:
            return []

        status, data = client.uid("search", None, "ALL")
        if status != "OK" or not data or not data[0]:
            return []
        uids = list(reversed(data[0].split()[-top:]))

        for uid in uids:
            uid_str = uid.decode("ascii", errors="ignore")
            f_status, raw_data = client.uid("fetch", uid, "(FLAGS RFC822)")
            if f_status != "OK" or not raw_data:
                continue
            raw_bytes = None
            flags_blob = b""
            for part in raw_data:
                if isinstance(part, tuple) and len(part) >= 2:
                    flags_blob += part[0] if isinstance(part[0], bytes) else b""
                    if isinstance(part[1], bytes):
                        raw_bytes = part[1]
            if not raw_bytes:
                continue

            parsed = email.message_from_bytes(raw_bytes)
            subject = _decode_mime_header(parsed.get("Subject")) or "(Без темы)"
            raw_from = _decode_mime_header(parsed.get("From"))
            sender_name, sender_email = parseaddr(raw_from)
            sender_name = sender_name.strip() or sender_email
            raw_to = _decode_mime_header(parsed.get("To"))
            _, to_addr = parseaddr(raw_to)
            received_at = _parse_iso_date(parsed.get("Date") or "")
            is_read = b"\\Seen" in flags_blob
            body_content, body_content_type, preview = _extract_body_rich(parsed)

            items.append(
                {
                    "id": f"imap:{folder}:{uid_str}",
                    "subject": subject,
                    "from": {
                        "emailAddress": {
                            "name": sender_name,
                            "address": sender_email,
                        }
                    },
                    "toRecipients": (
                        [{"emailAddress": {"address": to_addr}}] if to_addr else []
                    ),
                    "receivedDateTime": received_at,
                    "isRead": is_read,
                    "hasAttachments": False,
                    "bodyPreview": preview,
                    "bodyContent": body_content,
                    "bodyContentType": body_content_type,
                    "webLink": "",
                }
            )
    finally:
        try:
            client.logout()
        except Exception:
            pass

    return items
