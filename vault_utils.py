"""Utilities for parsing account lists and generating RFC 6238 TOTP 2FA codes."""

from __future__ import annotations

import base64
import hashlib
import hmac
import re
import struct
import time

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def generate_totp(
    secret: str,
    *,
    timestamp: float | None = None,
    step: int = 30,
    digits: int = 6,
) -> tuple[str, int]:
    """Return (6-digit TOTP code, seconds_remaining) for a Base32 secret."""
    cleaned = re.sub(r"[\s\-]+", "", (secret or "")).upper()
    if not cleaned:
        return "", 0
    padding = "=" * (-len(cleaned) % 8)
    try:
        key = base64.b32decode(cleaned + padding, casefold=True)
    except Exception:
        return "", 0

    now = time.time() if timestamp is None else timestamp
    counter = int(now // step)
    remaining = step - int(now % step)
    msg = struct.pack(">Q", counter)
    digest = hmac.new(key, msg, hashlib.sha1).digest()
    offset = digest[-1] & 0x0F
    code_int = struct.unpack(">I", digest[offset : offset + 4])[0] & 0x7FFFFFFF
    code = str(code_int % (10**digits)).zfill(digits)
    return code, remaining


def parse_account_lines(raw_text: str) -> tuple[list[dict[str, str]], list[str]]:
    """Parse pasted account lines in common colon-separated formats.

    Supported formats:
      - tt_login:tt_pass:outlook_email:outlook_pass[:[:]]recovery_email:recovery_pass[:totp]
      - outlook_email:outlook_pass[:recovery_email:recovery_pass[:totp]]
    Non-account lines (blank lines, Telegram headers, format descriptions) are skipped.
    """
    parsed: list[dict[str, str]] = []
    skipped: list[str] = []

    for raw_line in (raw_text or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        # Ignore Telegram headers / comments / format hints
        lowered = line.lower()
        if (
            line.startswith("[")
            or lowered.startswith("forwarded from")
            or lowered.startswith("формат")
            or "@" not in line
        ):
            continue

        # Collapse accidental double colons (e.g. "pass::recovery@mail")
        tokens = [part.strip() for part in line.split(":") if part.strip()]
        if len(tokens) < 2:
            skipped.append(line)
            continue

        email_indices = [i for i, tok in enumerate(tokens) if EMAIL_RE.match(tok)]
        if not email_indices:
            skipped.append(line)
            continue

        first_email_idx = email_indices[0]
        entry: dict[str, str] = {
            "tt_login": "",
            "tt_password": "",
            "outlook_email": "",
            "outlook_password": "",
            "recovery_email": "",
            "recovery_password": "",
            "totp_secret": "",
        }

        if first_email_idx == 2:
            # Format: tt_login : tt_pass : outlook_email : outlook_pass [: recovery_email : recovery_pass [: totp]]
            entry["tt_login"] = tokens[0]
            entry["tt_password"] = tokens[1]
            entry["outlook_email"] = tokens[2].lower()
            if len(tokens) > 3:
                entry["outlook_password"] = tokens[3]
            if len(tokens) > 4 and EMAIL_RE.match(tokens[4]):
                entry["recovery_email"] = tokens[4].lower()
                if len(tokens) > 5:
                    entry["recovery_password"] = tokens[5]
                if len(tokens) > 6:
                    entry["totp_secret"] = tokens[6]
            elif len(tokens) > 4:
                entry["totp_secret"] = tokens[4]
        elif first_email_idx == 0:
            # Format: outlook_email : outlook_pass [: recovery_email : recovery_pass [: totp]]
            entry["outlook_email"] = tokens[0].lower()
            entry["outlook_password"] = tokens[1]
            if len(tokens) > 2 and EMAIL_RE.match(tokens[2]):
                entry["recovery_email"] = tokens[2].lower()
                if len(tokens) > 3:
                    entry["recovery_password"] = tokens[3]
                if len(tokens) > 4:
                    entry["totp_secret"] = tokens[4]
            elif len(tokens) > 2:
                entry["totp_secret"] = tokens[2]
        else:
            # Generic fallback: anchor on first email found
            entry["outlook_email"] = tokens[first_email_idx].lower()
            if first_email_idx >= 2:
                entry["tt_login"] = tokens[0]
                entry["tt_password"] = tokens[1]
            if len(tokens) > first_email_idx + 1:
                entry["outlook_password"] = tokens[first_email_idx + 1]
            if len(tokens) > first_email_idx + 2 and EMAIL_RE.match(
                tokens[first_email_idx + 2]
            ):
                entry["recovery_email"] = tokens[first_email_idx + 2].lower()
                if len(tokens) > first_email_idx + 3:
                    entry["recovery_password"] = tokens[first_email_idx + 3]
                if len(tokens) > first_email_idx + 4:
                    entry["totp_secret"] = tokens[first_email_idx + 4]

        if entry["outlook_email"]:
            parsed.append(entry)
        else:
            skipped.append(line)

    return parsed, skipped
