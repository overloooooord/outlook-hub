"""Telegram Bot for automated Outlook OAuth2 onboarding and instant TikTok code retrieval.

Interactive KURDY-style mail code retrieval via Microsoft Graph API (works even when IMAP is disabled),
with strict per-Telegram-ID profile isolation, separate 1-click credential copying, and instant code retry.
"""

from __future__ import annotations

import asyncio
import html
import io
import logging
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import httpx

ROOT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT_DIR))


def _load_local_dotenv(env_path: Path) -> None:
    """Load key=value pairs from .env if present without external dependencies."""
    if not env_path.is_file():
        return
    try:
        for raw_line in env_path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, val = line.split("=", 1)
            key = key.strip()
            val = val.strip().strip('"').strip("'")
            if key:
                os.environ.setdefault(key, val)
    except OSError:
        pass


_load_local_dotenv(ROOT_DIR / ".env")

from db import Database
from vault_utils import parse_account_lines
from auto_login import auto_login_account
from imap_client import _html_to_text, extract_verification_codes

BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
# Personal App ID with Microsoft Graph Mail.Read (works on 100% of Outlook/Hotmail accounts,
# even when Microsoft disables IMAP on consumer mailboxes).
DEFAULT_APP_ID = os.environ.get(
    "OUTLOOK_HUB_CLIENT_ID", "a1c2e701-25ae-4029-b47e-655ea4919fea"
)
HUB_URL = os.environ.get("OUTLOOK_HUB_URL", "http://127.0.0.1:8765")
DB_PATH = ROOT_DIR / "data" / "outlook_hub.db"

# Consider verification emails received within the last 5 minutes as fresh codes.
FRESH_CODE_WINDOW_SEC = 300
# Default polling timeout when waiting for an incoming code.
CODE_POLL_TIMEOUT_SEC = 60
CODE_POLL_INTERVAL_SEC = 1.0

# Display timezone offset (UTC+3 / Moscow)
LOCAL_TZ = timezone(timedelta(hours=3))

# Automated Microsoft notification subjects/phrases that are never login OTPs.
_NOISE_PHRASES = (
    "new app(s) connected",
    "security info was added",
    "updates to our terms",
    "new login detected",
    "thanks for verifying",
    "новые приложения подключены",
    "обновление условий",
)

_OTP_KEYWORDS = (
    "code",
    "код",
    "verify",
    "verification",
    "otp",
    "confirm",
    "passcode",
    "security code",
    "подтвержд",
    "активац",
    "tiktok",
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("telegram_bot")


def _parse_iso_ts(iso_str: str) -> float:
    if not iso_str:
        return 0.0
    try:
        dt = datetime.fromisoformat(iso_str.replace("Z", "+00:00"))
        return dt.timestamp()
    except Exception:
        return 0.0


def _format_local_dt(ts: float) -> str:
    if ts <= 0:
        return ""
    return datetime.fromtimestamp(ts, tz=LOCAL_TZ).strftime("%d.%m %H:%M:%S")


def _get_profile_id(update_obj: dict[str, Any], fallback_chat_id: int) -> int:
    """Resolve isolated profile ID strictly from Telegram user ID (falls back to chat ID)."""
    sender = update_obj.get("from") or {}
    uid = sender.get("id")
    if uid is not None:
        return int(uid)
    return int(fallback_chat_id)


def _copy_btn(label: str, value: str) -> dict[str, Any] | None:
    """Build a Telegram Bot API 7.11+ native 1-click CopyTextButton."""
    val = str(value or "").strip()
    if not val or len(val) > 256:
        return None
    return {"text": label, "copy_text": {"text": val}}


def _format_cred_lines(
    *,
    tt_u: str = "",
    tt_p: str = "",
    email: str = "",
    outlook_pass: str = "",
) -> str:
    """Format credentials with separate <code> blocks so login, password, and email copy individually."""
    lines: list[str] = []
    if tt_u or tt_p:
        if tt_u:
            lines.append(f"👤 <b>Логин:</b> <code>{html.escape(tt_u)}</code>")
        if tt_p:
            lines.append(f"🔑 <b>Пароль:</b> <code>{html.escape(tt_p)}</code>")
        if email:
            lines.append(f"📧 <b>Почта:</b> <code>{html.escape(email)}</code>")
    else:
        if email:
            lines.append(f"📧 <b>Почта:</b> <code>{html.escape(email)}</code>")
        if outlook_pass:
            lines.append(f"🔑 <b>Пароль:</b> <code>{html.escape(outlook_pass)}</code>")
    return "\n".join(lines)


def _build_copy_rows(
    *,
    code: str = "",
    code_label: str = "📋 Скопировать код",
    tt_u: str = "",
    tt_p: str = "",
    email: str = "",
    outlook_pass: str = "",
) -> list[list[dict[str, Any]]]:
    """Build convenient 1-click copy inline keyboard rows with separate Login and Password buttons."""
    rows: list[list[dict[str, Any]]] = []
    if code:
        btn = _copy_btn(f"{code_label}: {code}", code)
        if btn:
            rows.append([btn])

    login_val = tt_u or email
    pass_val = tt_p or outlook_pass
    cred_row: list[dict[str, Any]] = []
    if login_val:
        b = _copy_btn("📋 Логин", login_val)
        if b:
            cred_row.append(b)
    if pass_val:
        b = _copy_btn("📋 Пароль", pass_val)
        if b:
            cred_row.append(b)
    if cred_row:
        rows.append(cred_row)
    return rows


def _extract_otp_from_message(msg: dict[str, Any]) -> tuple[str | None, bool]:
    """Extract OTP code from a normalized message dict.

    Returns (code, is_tiktok). Skips Microsoft app-connected / security alert noise.
    """
    subject = str(msg.get("subject") or "")
    sender = str(msg.get("sender") or "")
    preview = str(msg.get("preview") or "")
    body_text = str(msg.get("body_text") or preview)

    subj_low = subject.lower()
    combined_low = f"{sender} {subject} {preview}".lower()

    # Ignore Microsoft automated security/app notifications that don't contain an OTP
    if any(p in subj_low or p in preview.lower() for p in _NOISE_PHRASES):
        if "security code" not in subj_low and "код безопасности" not in subj_low:
            return None, False

    is_tiktok = "tiktok" in combined_low or "tiktok" in body_text.lower()

    # Priority 1: TikTok 6-digit code (usually directly in Subject: "510074 is your 6-digit code")
    if is_tiktok:
        m = re.search(r"\b(\d{6})\b", subject)
        if m:
            return m.group(1), True
        m = re.search(r"\b(\d{6})\b", preview)
        if m:
            return m.group(1), True
        m = re.search(r"\b(\d{6})\b", body_text)
        if m:
            return m.group(1), True

    # Priority 2: Other verification/OTP emails with explicit code keywords
    if any(kw in combined_low for kw in _OTP_KEYWORDS):
        codes = extract_verification_codes(subject, body_text or preview)
        if codes:
            return codes[0], is_tiktok

    return None, False


class TelegramBot:
    def __init__(self, token: str, hub_url: str = HUB_URL, db_path: Path = DB_PATH) -> None:
        self.token = token
        self.api_base = f"https://api.telegram.org/bot{token}"
        self.hub_url = hub_url
        self.db = Database(db_path)
        self.http: httpx.AsyncClient | None = None
        self.running = True
        # Scoped strictly per (profile_id, vault_id) so actions in one TG account never affect another.
        self.active_searches: set[tuple[int, int]] = set()
        self.cancelled_searches: set[tuple[int, int]] = set()
        # Remember last delivered (msg_id, received_ts, code) per (profile_id, vault_id) for "Повторить".
        self.last_delivered_msg: dict[tuple[int, int], tuple[str, float, str]] = {}
        # Limit concurrent headless Chrome instances during background warmup / re-auth
        self.login_semaphore = asyncio.Semaphore(1)

    def _cancel_profile_searches(self, profile_id: int) -> None:
        """Cancel active code-polling tasks belonging only to this Telegram profile."""
        pid = int(profile_id)
        for key in list(self.active_searches):
            if key[0] == pid:
                self.cancelled_searches.add(key)

    async def api_call(
        self,
        method: str,
        data: dict[str, Any] | None = None,
        files: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        url = f"{self.api_base}/{method}"
        for attempt in range(3):
            try:
                assert self.http is not None
                if files:
                    resp = await self.http.post(url, data=data, files=files, timeout=30.0)
                else:
                    resp = await self.http.post(url, json=data or {}, timeout=30.0)
                payload = resp.json()
                if not payload.get("ok"):
                    desc = str(payload.get("description") or "")
                    if "message is not modified" not in desc.lower():
                        logger.warning(f"Telegram API {method} returned ok=False: {desc}")
                return payload
            except Exception as exc:
                if attempt == 2:
                    logger.warning(f"Telegram API call {method} failed: {exc}")
                    raise
                await asyncio.sleep(1.0)
        return {}

    async def send_message(
        self,
        chat_id: int | str,
        text: str,
        parse_mode: str | None = None,
        reply_to_message_id: int | None = None,
        reply_markup: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        data: dict[str, Any] = {"chat_id": chat_id, "text": text}
        if parse_mode:
            data["parse_mode"] = parse_mode
        if reply_to_message_id:
            data["reply_to_message_id"] = reply_to_message_id
        if reply_markup:
            data["reply_markup"] = reply_markup
        return await self.api_call("sendMessage", data=data)

    async def edit_message(
        self,
        chat_id: int | str,
        message_id: int,
        text: str,
        parse_mode: str | None = None,
        reply_markup: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        data: dict[str, Any] = {
            "chat_id": chat_id,
            "message_id": message_id,
            "text": text,
        }
        if parse_mode:
            data["parse_mode"] = parse_mode
        if reply_markup:
            data["reply_markup"] = reply_markup
        return await self.api_call("editMessageText", data=data)

    async def answer_callback_query(
        self, callback_query_id: str, text: str = "", show_alert: bool = False
    ) -> dict[str, Any]:
        return await self.api_call(
            "answerCallbackQuery",
            {"callback_query_id": callback_query_id, "text": text, "show_alert": show_alert},
        )

    async def send_document(
        self,
        chat_id: int | str,
        filename: str,
        content: str,
        caption: str = "",
    ) -> dict[str, Any]:
        files = {"document": (filename, io.BytesIO(content.encode("utf-8")), "text/plain")}
        data: dict[str, Any] = {"chat_id": chat_id}
        if caption:
            data["caption"] = caption
        return await self.api_call("sendDocument", data=data, files=files)

    def get_current_client_id(self) -> str:
        s = self.db.get_settings()
        return s.get("client_id") or DEFAULT_APP_ID

    def build_mailboxes_keyboard(
        self, accounts: list[dict[str, Any]], page: int = 0, per_page: int = 5
    ) -> dict[str, Any]:
        total = len(accounts)
        total_pages = max(1, (total + per_page - 1) // per_page)
        page = max(0, min(page, total_pages - 1))

        start = page * per_page
        end = start + per_page
        page_items = accounts[start:end]

        keyboard: list[list[dict[str, str]]] = []
        for acc in page_items:
            email = acc.get("outlook_email") or ""
            keyboard.append([{"text": email, "callback_data": f"mail:{acc['id']}"}])

        # Pagination row (KURDY style)
        if total_pages > 1:
            nav_row: list[dict[str, str]] = []
            if page > 0:
                nav_row.append({"text": "◀", "callback_data": f"page:{page - 1}"})
            nav_row.append({"text": f"{page + 1}/{total_pages}", "callback_data": "noop"})
            if page < total_pages - 1:
                nav_row.append({"text": "▶", "callback_data": f"page:{page + 1}"})
                if page < total_pages - 2:
                    nav_row.append({"text": "⏩", "callback_data": f"page:{total_pages - 1}"})
            keyboard.append(nav_row)

        # Control buttons rows
        keyboard.append([
            {"text": "🔄 Обновить", "callback_data": f"page:{page}"},
            {"text": "📥 Экспорт .txt", "callback_data": "export"},
        ])
        keyboard.append([
            {"text": f"🗑 Удалить все почты ({total})", "callback_data": "clear_all_ask"},
        ])

        return {"inline_keyboard": keyboard}

    async def show_mailboxes_menu(
        self,
        chat_id: int,
        page: int = 0,
        edit_msg_id: int | None = None,
        profile_id: int | None = None,
    ) -> None:
        pid = int(chat_id if profile_id is None else profile_id)
        accounts = self.db.get_active_vault_accounts(pid)
        if not accounts:
            text = (
                "📭 <b>Список почт пуст</b>\n"
                f"👤 Профиль: <code>{pid}</code>\n\n"
                "Чтобы добавить почты в этот профиль, просто <b>отправьте сюда строки</b> или .txt файл:\n"
                "<code>tt_user:tt_pass:box@outlook.com:pass:rec@mail.com:rec_pass:TOTP</code>"
            )
            if edit_msg_id:
                res = await self.edit_message(chat_id, edit_msg_id, text, parse_mode="HTML")
                if not res.get("ok") and "not modified" not in str(res.get("description", "")).lower():
                    await self.send_message(chat_id, text, parse_mode="HTML")
            else:
                await self.send_message(chat_id, text, parse_mode="HTML")
            return

        text = (
            "📢 <b>Выберите почту для получения кода от TikTok:</b>\n"
            "<i>! Чтобы добавить почты, отправьте боту данные в доступных форматах.</i>\n\n"
            f"👤 Профиль: <code>{pid}</code>\n"
            f"📬 Доступно почт: <b>{len(accounts)}</b>"
        )
        markup = self.build_mailboxes_keyboard(accounts, page=page)
        if edit_msg_id:
            res = await self.edit_message(
                chat_id, edit_msg_id, text, parse_mode="HTML", reply_markup=markup
            )
            if not res.get("ok") and "not modified" not in str(res.get("description", "")).lower():
                await self.send_message(chat_id, text, parse_mode="HTML", reply_markup=markup)
        else:
            await self.send_message(chat_id, text, parse_mode="HTML", reply_markup=markup)

    async def _refresh_account_token(self, acc: dict[str, Any]) -> tuple[str | None, str]:
        """Refresh OAuth2 access token using the account's own client_id."""
        rt = str(acc.get("refresh_token") or "").strip()
        if not rt:
            return None, "empty_refresh_token"

        client_id = str(acc.get("client_id") or "").strip() or self.get_current_client_id()
        protocol = str(acc.get("protocol") or "graph").strip().lower()
        tenant = str(acc.get("tenant") or "consumers").strip() or "consumers"
        if tenant == "common":
            tenant = "consumers"

        assert self.http is not None
        try:
            resp = await self.http.post(
                f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
                data={
                    "client_id": client_id,
                    "grant_type": "refresh_token",
                    "refresh_token": rt,
                },
                timeout=15.0,
            )
            if resp.status_code != 200:
                logger.warning(
                    f"Token refresh failed for {acc.get('email')} (cid={client_id[:8]}): "
                    f"HTTP {resp.status_code} {resp.text[:160]}"
                )
                return None, f"http_{resp.status_code}"

            tok = resp.json()
            at = str(tok.get("access_token") or "")
            new_rt = str(tok.get("refresh_token") or rt)
            expires_at = time.time() + int(tok.get("expires_in") or 3600)
            if at and acc.get("id"):
                self.db.update_account_tokens(
                    int(acc["id"]),
                    access_token=at,
                    refresh_token=new_rt,
                    token_expires_at=expires_at,
                )
            return at, protocol
        except Exception as exc:
            logger.warning(f"Token refresh error for {acc.get('email')}: {exc}")
            return None, str(exc)

    async def _fetch_messages_once(
        self, access_token: str, protocol: str = "graph", top: int = 6
    ) -> tuple[str, list[dict[str, Any]]]:
        """Fetch latest messages via Microsoft Graph API (or Outlook REST fallback).

        Lightweight query (no full HTML body download) -> ~350-500ms latency!
        Returns (status, messages) where status is 'ok' | 'unauthorized' | 'error'.
        """
        assert self.http is not None
        headers = {"Authorization": f"Bearer {access_token}"}

        # 1. Microsoft Graph API (works across Inbox + Junk even when IMAP is disabled!)
        if protocol == "graph":
            try:
                url = (
                    f"https://graph.microsoft.com/v1.0/me/messages"
                    f"?$top={top}&$orderby=receivedDateTime desc"
                    f"&$select=id,subject,from,receivedDateTime,bodyPreview"
                )
                resp = await self.http.get(url, headers=headers, timeout=10.0)
                if resp.status_code == 200:
                    raw = resp.json().get("value") or []
                    msgs: list[dict[str, Any]] = []
                    for m in raw:
                        sender_obj = (m.get("from") or {}).get("emailAddress") or {}
                        sender = str(sender_obj.get("address") or sender_obj.get("name") or "")
                        rec_iso = str(m.get("receivedDateTime") or "")
                        prev = str(m.get("bodyPreview") or "")
                        msgs.append({
                            "id": str(m.get("id") or ""),
                            "subject": str(m.get("subject") or ""),
                            "sender": sender,
                            "received_at": rec_iso,
                            "received_ts": _parse_iso_ts(rec_iso),
                            "preview": prev,
                            "body_text": prev,
                        })
                    return "ok", msgs
                if resp.status_code in (401, 403):
                    return "unauthorized", []
                logger.warning(f"Graph API error HTTP {resp.status_code}: {resp.text[:160]}")
                return "error", []
            except Exception as exc:
                logger.warning(f"Graph API fetch exception: {exc}")
                return "error", []

        # 2. Outlook REST API fallback (for Thunderbird tokens with Mail.Read)
        try:
            url = (
                f"https://outlook.office.com/api/v2.0/me/messages"
                f"?$top={top}&$orderby=ReceivedDateTime desc"
                f"&$select=Id,Subject,From,ReceivedDateTime,BodyPreview"
            )
            resp = await self.http.get(url, headers=headers, timeout=10.0)
            if resp.status_code == 200:
                raw = resp.json().get("value") or []
                msgs = []
                for m in raw:
                    sender_obj = (m.get("From") or {}).get("EmailAddress") or {}
                    sender = str(sender_obj.get("Address") or sender_obj.get("Name") or "")
                    rec_iso = str(m.get("ReceivedDateTime") or "")
                    prev = str(m.get("BodyPreview") or "")
                    msgs.append({
                        "id": str(m.get("Id") or ""),
                        "subject": str(m.get("Subject") or ""),
                        "sender": sender,
                        "received_at": rec_iso,
                        "received_ts": _parse_iso_ts(rec_iso),
                        "preview": prev,
                        "body_text": prev,
                    })
                return "ok", msgs
            if resp.status_code in (401, 403):
                return "unauthorized", []
            return "error", []
        except Exception as exc:
            logger.warning(f"Outlook REST fetch exception: {exc}")
            return "error", []

    async def _ensure_working_access(
        self, chat_id: int, entry: dict[str, Any], top: int = 6
    ) -> tuple[str | None, str, list[dict[str, Any]], str]:
        """Ensure account has a live Graph API token AND fetch initial messages in 1 roundtrip.

        Returns (access_token, protocol, initial_messages, error_message).
        """
        email = entry["outlook_email"]
        acc = self.db.get_account_by_email(email, include_tokens=True)

        # Fast path: 1 single HTTP request using cached access_token!
        if acc and acc.get("refresh_token"):
            cached_at = str(acc.get("access_token") or "")
            exp = float(acc.get("token_expires_at") or 0.0)
            proto = str(acc.get("protocol") or "graph").lower()
            if cached_at and exp > time.time() + 90 and proto == "graph":
                status, msgs = await self._fetch_messages_once(cached_at, protocol=proto, top=top)
                if status == "ok":
                    return cached_at, proto, msgs, ""

            # Refresh token once and fetch messages
            at, proto = await self._refresh_account_token(acc)
            if at:
                status, msgs = await self._fetch_messages_once(at, protocol=proto, top=top)
                if status == "ok":
                    return at, proto, msgs, ""
                logger.info(
                    f"Existing token for {email} returned {status} (proto={proto}); "
                    "re-authorizing with Graph API..."
                )

        # Slow path: auto-login via headless browser with Graph API (DEFAULT_APP_ID)
        await self.send_message(
            chat_id,
            f"⚙️ Авторизую почту <code>{html.escape(email)}</code> (~15 сек)...",
            parse_mode="HTML",
        )
        self.db.update_settings({"client_id": DEFAULT_APP_ID})
        try:
            async with self.login_semaphore:
                await asyncio.to_thread(
                    auto_login_account, entry, base_url=self.hub_url, headless=True
                )
        except Exception as err:
            return None, "graph", [], str(err)

        acc = self.db.get_account_by_email(email, include_tokens=True)
        if not acc or not acc.get("refresh_token"):
            return None, "graph", [], "Токен не найден после авторизации."

        at = str(acc.get("access_token") or "")
        proto = str(acc.get("protocol") or "graph").lower()
        if not at:
            at, proto = await self._refresh_account_token(acc)
        if not at:
            return None, proto, [], "Не удалось получить access_token."
        status, msgs = await self._fetch_messages_once(at, protocol=proto, top=top)
        return at, proto, (msgs if status == "ok" else []), ""

    async def request_code_for_vault(
        self,
        chat_id: int,
        vault_id: int,
        profile_id: int | None = None,
        skip_prev_code: str = "",
    ) -> None:
        pid = int(chat_id if profile_id is None else profile_id)
        search_key = (pid, vault_id)
        if search_key in self.active_searches:
            await self.send_message(chat_id, "⏳ Поиск кода для этой почты уже идёт...")
            return

        entry = self.db.get_vault_entry(vault_id, chat_id=pid)
        if not entry:
            await self.send_message(chat_id, "❌ Почта не найдена в этом профиле или уже удалена.")
            return

        self.active_searches.add(search_key)
        self.cancelled_searches.discard(search_key)
        email = entry["outlook_email"]
        outlook_pass = entry.get("outlook_password") or ""
        tt_u = entry.get("tt_login") or ""
        tt_p = entry.get("tt_password") or ""
        was_active = str(entry.get("status") or "active") == "active"
        cred_block = _format_cred_lines(
            tt_u=tt_u, tt_p=tt_p, email=email, outlook_pass=outlook_pass
        )
        click_ts = time.time()
        is_retry = bool(skip_prev_code)
        prev_delivered = self.last_delivered_msg.get(search_key) if is_retry else None
        retry_cb = (
            f"retry:{vault_id}:{skip_prev_code[:16]}" if skip_prev_code else f"mail:{vault_id}"
        )

        try:
            init_header = (
                "🔁 <b>Повторный поиск нового кода...</b>"
                if is_retry
                else "📢 <b>Начинаю поиск кода...</b>"
            )
            init_resp = await self.send_message(
                chat_id,
                f"{init_header}\n{cred_block}",
                parse_mode="HTML",
            )
            status_msg_id = (init_resp.get("result") or {}).get("message_id")

            # Single fast call (~0.4s) returns BOTH live token AND initial messages!
            access_token, proto, msgs, err = await self._ensure_working_access(chat_id, entry, top=6)
            if not access_token:
                fail_rows = _build_copy_rows(
                    tt_u=tt_u, tt_p=tt_p, email=email, outlook_pass=outlook_pass
                )
                fail_rows.extend([
                    [{"text": "🔁 Повторить", "callback_data": retry_cb}],
                    [
                        {"text": "🗑 Удалить почту", "callback_data": f"del:{vault_id}"},
                        {"text": "🔙 Список почт", "callback_data": "menu_mails"},
                    ],
                ])
                fail_kb = {"inline_keyboard": fail_rows}
                await self.send_message(
                    chat_id,
                    f"❌ Не удалось подключиться к <code>{html.escape(email)}</code>:\n"
                    f"<code>{html.escape(err[:300])}</code>",
                    parse_mode="HTML",
                    reply_markup=fail_kb,
                )
                return

            def _scan_messages(
                msg_list: list[dict[str, Any]], cutoff: float
            ) -> tuple[tuple[str, str, float] | None, tuple[str, float, str] | None]:
                fresh_candidates: list[tuple[bool, float, str, str]] = []
                old_candidates: list[tuple[bool, float, str, str]] = []
                for m in msg_list:
                    code_val, is_tiktok = _extract_otp_from_message(m)
                    if not code_val:
                        continue
                    rec_ts = float(m.get("received_ts") or 0.0)
                    mid = str(m.get("id") or "")
                    if is_retry:
                        if prev_delivered is not None:
                            prev_mid, prev_ts, _ = prev_delivered
                            is_fresh = (mid != prev_mid) and (rec_ts > prev_ts)
                        else:
                            is_fresh = (code_val != skip_prev_code) and (rec_ts >= click_ts - 30)
                    else:
                        is_fresh = rec_ts >= cutoff

                    if is_fresh:
                        fresh_candidates.append((is_tiktok, rec_ts, code_val, mid))
                    else:
                        old_candidates.append((is_tiktok, rec_ts, code_val, str(m.get("subject") or "")))

                old_best: tuple[str, float, str] | None = None
                if old_candidates:
                    old_candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
                    old_best = (old_candidates[0][2], old_candidates[0][1], old_candidates[0][3])
                if fresh_candidates:
                    fresh_candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
                    top_c = fresh_candidates[0]
                    return (top_c[2], top_c[3], top_c[1]), old_best
                return None, old_best

            fresh_cutoff = click_ts - FRESH_CODE_WINDOW_SEC
            found_hit, last_old_code = _scan_messages(msgs, fresh_cutoff)

            # If NO fresh code on instant check (<0.5s), IMMEDIATELY show last code in inbox
            # + separate 1-click copy buttons for login/pass/old code + live waiting indicator!
            if not found_hit and status_msg_id:
                old_line = ""
                old_code_val = ""
                if last_old_code:
                    oc, ots, _ = last_old_code
                    old_code_val = oc
                    old_label_txt = "Предыдущий код" if is_retry else "Последний код в ящике"
                    old_line = (
                        f"\n📬 <b>{old_label_txt}:</b> <code>{html.escape(oc)}</code> "
                        f"<i>({_format_local_dt(ots)} МСК)</i>"
                    )
                wait_rows = _build_copy_rows(
                    code=old_code_val,
                    code_label="📋 Старый код",
                    tt_u=tt_u,
                    tt_p=tt_p,
                    email=email,
                    outlook_pass=outlook_pass,
                )
                wait_rows.append([{"text": "📨 Показать письма", "callback_data": f"inbox:{vault_id}"}])
                wait_rows.append([
                    {"text": "🗑 Удалить почту", "callback_data": f"del:{vault_id}"},
                    {"text": "❌ Стоп", "callback_data": f"stop:{vault_id}"},
                ])
                wait_kb = {"inline_keyboard": wait_rows}
                await self.edit_message(
                    chat_id,
                    status_msg_id,
                    f"📢 <b>Жду новый код...</b>\n"
                    f"{cred_block}"
                    f"{old_line}\n\n"
                    f"⚡ <i>Как только TikTok пришлёт новый код — сразу выдам сюда (проверяю каждую 1 сек).</i>",
                    parse_mode="HTML",
                    reply_markup=wait_kb,
                )

            deadline = time.time() + CODE_POLL_TIMEOUT_SEC
            while not found_hit and time.time() < deadline:
                if search_key in self.cancelled_searches:
                    if status_msg_id:
                        stop_rows = _build_copy_rows(
                            tt_u=tt_u, tt_p=tt_p, email=email, outlook_pass=outlook_pass
                        )
                        stop_rows.append([
                            {"text": "🔁 Повторить", "callback_data": retry_cb},
                            {"text": "🔙 Список почт", "callback_data": "menu_mails"},
                        ])
                        await self.edit_message(
                            chat_id,
                            status_msg_id,
                            f"⏹ Поиск кода остановлен.\n{cred_block}",
                            parse_mode="HTML",
                            reply_markup={"inline_keyboard": stop_rows},
                        )
                    return

                await asyncio.sleep(CODE_POLL_INTERVAL_SEC)
                if search_key in self.cancelled_searches:
                    if status_msg_id:
                        stop_rows = _build_copy_rows(
                            tt_u=tt_u, tt_p=tt_p, email=email, outlook_pass=outlook_pass
                        )
                        stop_rows.append([
                            {"text": "🔁 Повторить", "callback_data": retry_cb},
                            {"text": "🔙 Список почт", "callback_data": "menu_mails"},
                        ])
                        await self.edit_message(
                            chat_id,
                            status_msg_id,
                            f"⏹ Поиск кода остановлен.\n{cred_block}",
                            parse_mode="HTML",
                            reply_markup={"inline_keyboard": stop_rows},
                        )
                    return

                status, poll_msgs = await self._fetch_messages_once(access_token, protocol=proto, top=6)
                if status == "unauthorized":
                    acc = self.db.get_account_by_email(email, include_tokens=True)
                    if acc:
                        new_at, proto = await self._refresh_account_token(acc)
                        if new_at:
                            access_token = new_at
                            status, poll_msgs = await self._fetch_messages_once(
                                access_token, protocol=proto, top=6
                            )
                    if status == "unauthorized":
                        break

                if status == "ok":
                    hit, oc = _scan_messages(poll_msgs, fresh_cutoff)
                    if oc and not last_old_code:
                        last_old_code = oc
                    if hit:
                        found_hit = hit
                        break

            if found_hit:
                found_code, found_mid, found_ts = found_hit
                self.last_delivered_msg[search_key] = (found_mid, found_ts, found_code)
                # Mark as used -> removed from this profile's "Мои почты"!
                self.db.mark_vault_account_used(vault_id, chat_id=pid)

                if status_msg_id:
                    await self.edit_message(
                        chat_id,
                        status_msg_id,
                        f"✅ Код получен для <code>{html.escape(email)}</code>",
                        parse_mode="HTML",
                    )

                success_text = (
                    f"📢 <b>Код с почты:</b> <code>{html.escape(found_code)}</code>\n\n"
                    f"{cred_block}"
                )
                copy_rows = _build_copy_rows(
                    code=found_code,
                    code_label="📋 Скопировать код",
                    tt_u=tt_u,
                    tt_p=tt_p,
                    email=email,
                    outlook_pass=outlook_pass,
                )
                next_retry_cb = f"retry:{vault_id}:{found_code[:16]}"
                copy_rows.append([
                    {"text": "🔁 Повторить", "callback_data": next_retry_cb},
                    {"text": "📨 Письма", "callback_data": f"inbox:{vault_id}"},
                ])
                if not was_active:
                    copy_rows.append([{"text": "🔙 Список почт", "callback_data": "menu_mails"}])
                success_kb = {"inline_keyboard": copy_rows}

                # Release search lock before sending menu so instant "Повторить" click works immediately.
                self.active_searches.discard(search_key)
                self.cancelled_searches.discard(search_key)

                await self.send_message(
                    chat_id, success_text, parse_mode="HTML", reply_markup=success_kb
                )

                # Automatically show updated remaining mailboxes on first code from active mailbox
                if was_active:
                    await self.show_mailboxes_menu(chat_id, profile_id=pid)
            else:
                old_hint = ""
                old_code_val = ""
                if last_old_code:
                    oc, ots, _ = last_old_code
                    old_code_val = oc
                    old_hint = (
                        f"\n\n📌 Последний <b>старый</b> код в ящике: "
                        f"<code>{html.escape(oc)}</code> ({_format_local_dt(ots)} МСК)"
                    )

                fail_rows = _build_copy_rows(
                    code=old_code_val,
                    code_label="📋 Скопировать код",
                    tt_u=tt_u,
                    tt_p=tt_p,
                    email=email,
                    outlook_pass=outlook_pass,
                )
                fail_rows.extend([
                    [{"text": "🔁 Повторить", "callback_data": retry_cb}],
                    [{"text": "📨 Показать последние письма", "callback_data": f"inbox:{vault_id}"}],
                    [
                        {"text": "🗑 Удалить почту", "callback_data": f"del:{vault_id}"},
                        {"text": "🔙 Список почт", "callback_data": "menu_mails"},
                    ],
                ])
                fail_kb = {"inline_keyboard": fail_rows}
                timeout_text = (
                    f"⚠️ <b>Новый код не пришел за {CODE_POLL_TIMEOUT_SEC} сек.</b>\n"
                    f"{cred_block}"
                    f"{old_hint}\n\n"
                    f"Отправьте код в TikTok и нажмите «🔁 Повторить»."
                )
                if status_msg_id:
                    await self.edit_message(
                        chat_id, status_msg_id, timeout_text, parse_mode="HTML", reply_markup=fail_kb
                    )
                else:
                    await self.send_message(
                        chat_id, timeout_text, parse_mode="HTML", reply_markup=fail_kb
                    )

        finally:
            self.active_searches.discard(search_key)
            self.cancelled_searches.discard(search_key)

    async def show_mailbox_inbox(
        self, chat_id: int, vault_id: int, profile_id: int | None = None
    ) -> None:
        """Show the latest 5 emails in a mailbox right inside Telegram (1 fast Graph call)."""
        pid = int(chat_id if profile_id is None else profile_id)
        entry = self.db.get_vault_entry(vault_id, chat_id=pid)
        if not entry:
            await self.send_message(chat_id, "❌ Почта не найдена в этом профиле.")
            return

        email = entry["outlook_email"]
        outlook_pass = entry.get("outlook_password") or ""
        tt_u = entry.get("tt_login") or ""
        tt_p = entry.get("tt_password") or ""
        access_token, proto, msgs, err = await self._ensure_working_access(chat_id, entry, top=5)
        if not access_token:
            await self.send_message(
                chat_id,
                f"❌ Ошибка доступа к <code>{html.escape(email)}</code>: {html.escape(err[:200])}",
                parse_mode="HTML",
            )
            return

        if not msgs:
            await self.send_message(
                chat_id,
                f"📭 Для <code>{html.escape(email)}</code> не найдено писем.",
                parse_mode="HTML",
            )
            return

        blocks = [f"📬 <b>Последние письма для {html.escape(email)}:</b>"]
        newest_code = ""
        for m in msgs[:5]:
            code_val, _ = _extract_otp_from_message(m)
            if code_val and not newest_code:
                newest_code = code_val
            code_line = f"🔢 <b>КОД:</b> <code>{html.escape(code_val)}</code>\n" if code_val else ""
            dt_str = _format_local_dt(float(m.get("received_ts") or 0.0))
            subj = html.escape(str(m.get("subject") or "(без темы)")[:90])
            sender = html.escape(str(m.get("sender") or "")[:60])
            prev = html.escape(str(m.get("preview") or "").replace("\r", " ").replace("\n", " ")[:110])
            blocks.append(
                f"{code_line}"
                f"📝 <b>{subj}</b>\n"
                f"👤 {sender} | ⏰ {dt_str}\n"
                f"<i>{prev}</i>"
            )

        kb_rows = _build_copy_rows(
            code=newest_code,
            code_label="📋 Скопировать код",
            tt_u=tt_u,
            tt_p=tt_p,
            email=email,
            outlook_pass=outlook_pass,
        )
        kb_rows.extend([
            [{"text": "🔁 Повторить", "callback_data": f"mail:{vault_id}"}],
            [
                {"text": "🗑 Удалить почту", "callback_data": f"del:{vault_id}"},
                {"text": "🔙 Список почт", "callback_data": "menu_mails"},
            ],
        ])
        kb = {"inline_keyboard": kb_rows}
        await self.send_message(chat_id, "\n\n".join(blocks), parse_mode="HTML", reply_markup=kb)

    async def export_accounts_txt(
        self, chat_id: int, profile_id: int | None = None
    ) -> None:
        pid = int(chat_id if profile_id is None else profile_id)
        accounts = self.db.get_active_vault_accounts(pid)
        if not accounts:
            accounts = [
                a for a in self.db.list_vault_entries(chat_id=pid)
                if a.get("status") != "deleted"
            ]
        if not accounts:
            await self.send_message(
                chat_id,
                f"📭 В профиле <code>{pid}</code> нет сохранённых аккаунтов.",
                parse_mode="HTML",
            )
            return

        lines: list[str] = []
        for acc in accounts:
            email = acc.get("outlook_email") or ""
            pwd = acc.get("outlook_password") or ""
            tt_u = acc.get("tt_login") or ""
            tt_p = acc.get("tt_password") or ""
            o_acc = self.db.get_account_by_email(email, include_tokens=True)
            rt = o_acc.get("refresh_token") if o_acc else ""
            cid = (o_acc.get("client_id") if o_acc else "") or self.get_current_client_id()
            if rt:
                lines.append(f"{email}|{pwd}|{tt_u}|{tt_p}|{rt}|{cid}")

        if not lines:
            await self.send_message(
                chat_id,
                f"⚠️ В профиле <code>{pid}</code> пока нет аккаунтов с полученным токеном.",
                parse_mode="HTML",
            )
            return

        doc = "\n".join(lines)
        await self.send_document(
            chat_id,
            filename=f"outlook_tokens_{pid}.txt",
            content=doc,
            caption=f"📄 Профиль {pid}: {len(lines)} аккаунтов с рабочими токенами",
        )

    async def handle_callback_query(self, cq: dict[str, Any]) -> None:
        cq_id = cq["id"]
        data = str(cq.get("data") or "")
        msg = cq.get("message") or {}
        chat_id = msg.get("chat", {}).get("id")
        msg_id = msg.get("message_id")

        if not chat_id:
            return
        profile_id = _get_profile_id(cq, chat_id)

        if data == "noop":
            await self.answer_callback_query(cq_id)
            return

        if data.startswith("page:"):
            page_num = int(data.split(":")[1])
            await self.answer_callback_query(cq_id)
            await self.show_mailboxes_menu(
                chat_id, page=page_num, edit_msg_id=msg_id, profile_id=profile_id
            )
            return

        if data == "menu_mails" or data == "refresh":
            await self.answer_callback_query(cq_id, "Обновлено")
            await self.show_mailboxes_menu(
                chat_id, page=0, edit_msg_id=msg_id, profile_id=profile_id
            )
            return

        if data == "clear_all_ask":
            active = self.db.get_active_vault_accounts(profile_id)
            count = len(active)
            await self.answer_callback_query(cq_id)
            if count == 0:
                await self.show_mailboxes_menu(
                    chat_id, page=0, edit_msg_id=msg_id, profile_id=profile_id
                )
                return
            confirm_kb = {
                "inline_keyboard": [
                    [{"text": f"✅ Да, удалить все почты ({count})", "callback_data": "clear_all_confirm"}],
                    [{"text": "❌ Отмена", "callback_data": "menu_mails"}],
                ]
            }
            await self.edit_message(
                chat_id,
                msg_id,
                f"⚠️ <b>Удалить все активные почты ({count} шт.) в этом профиле?</b>\n"
                f"👤 Профиль: <code>{profile_id}</code>\n"
                f"<i>Другие Telegram-профили не будут затронуты. Токены в базе сохранятся.</i>",
                parse_mode="HTML",
                reply_markup=confirm_kb,
            )
            return

        if data in ("clear_all", "clear_all_confirm"):
            self._cancel_profile_searches(profile_id)
            count = self.db.clear_active_vault_accounts(profile_id)
            await self.answer_callback_query(cq_id, f"Удалено почт: {count}")
            await self.show_mailboxes_menu(
                chat_id, page=0, edit_msg_id=msg_id, profile_id=profile_id
            )
            return

        if data.startswith("del:"):
            v_id = int(data.split(":")[1])
            self.cancelled_searches.add((profile_id, v_id))
            self.db.mark_vault_account_used(v_id, chat_id=profile_id)
            await self.answer_callback_query(cq_id, "Почта удалена")
            await self.show_mailboxes_menu(
                chat_id, page=0, edit_msg_id=msg_id, profile_id=profile_id
            )
            return

        if data.startswith("stop:"):
            v_id = int(data.split(":")[1])
            self.cancelled_searches.add((profile_id, v_id))
            await self.answer_callback_query(cq_id, "Поиск остановлен")
            return

        if data == "export":
            await self.answer_callback_query(cq_id, "Формирую файл...")
            await self.export_accounts_txt(chat_id, profile_id=profile_id)
            return

        if data.startswith("inbox:"):
            vault_id = int(data.split(":")[1])
            await self.answer_callback_query(cq_id, "Загружаю письма...")
            asyncio.create_task(
                self.show_mailbox_inbox(chat_id, vault_id, profile_id=profile_id)
            )
            return

        if data.startswith("retry:"):
            parts = data.split(":", 2)
            vault_id = int(parts[1])
            prev_code = parts[2] if len(parts) > 2 else ""
            await self.answer_callback_query(cq_id, "Ищу новый код...")
            asyncio.create_task(
                self.request_code_for_vault(
                    chat_id, vault_id, profile_id=profile_id, skip_prev_code=prev_code
                )
            )
            return

        if data.startswith("mail:"):
            vault_id = int(data.split(":")[1])
            await self.answer_callback_query(cq_id, "Ищу код...")
            asyncio.create_task(
                self.request_code_for_vault(chat_id, vault_id, profile_id=profile_id)
            )
            return

    async def background_warmup(
        self, profile_id: int, parsed_items: list[dict[str, Any]]
    ) -> None:
        """Background worker to auto-authorize new accounts via Graph API so codes fetch in <1s."""
        self.db.update_settings({"client_id": DEFAULT_APP_ID})
        for item in parsed_items:
            email = item.get("outlook_email")
            if not email:
                continue
            # Stop warming up accounts that were removed from this profile
            active_emails = {
                a["outlook_email"].lower() for a in self.db.get_active_vault_accounts(profile_id)
            }
            if email.lower() not in active_emails:
                continue
            acc = self.db.get_account_by_email(email, include_tokens=True)
            if acc and acc.get("refresh_token") and acc.get("protocol") == "graph":
                continue
            try:
                async with self.login_semaphore:
                    await asyncio.to_thread(
                        auto_login_account, item, base_url=self.hub_url, headless=True
                    )
                logger.info(f"Background Graph warmup complete for {email} (profile={profile_id})")
                await asyncio.sleep(2.5)
            except Exception as e:
                logger.warning(f"Background warmup failed for {email}: {e}")
                await asyncio.sleep(2.0)

    async def _download_document_text(self, document: dict[str, Any]) -> str:
        """Download a .txt document uploaded by the user in Telegram."""
        file_id = str(document.get("file_id") or "")
        if not file_id:
            return ""
        info = await self.api_call("getFile", data={"file_id": file_id})
        file_path = str((info.get("result") or {}).get("file_path") or "")
        if not file_path or self.http is None:
            return ""
        dl_url = f"https://api.telegram.org/file/bot{self.token}/{file_path}"
        resp = await self.http.get(dl_url, timeout=20.0)
        if resp.status_code != 200:
            return ""
        return resp.content.decode("utf-8", errors="replace")

    async def handle_message(self, message: dict[str, Any]) -> None:
        chat_id = message.get("chat", {}).get("id")
        text = str(message.get("text") or "").strip()
        if not chat_id:
            return
        profile_id = _get_profile_id(message, chat_id)
        sender_info = message.get("from") or {}
        username = str(sender_info.get("username") or "").strip()

        # Allow uploading a .txt file with accounts directly
        if not text and message.get("document"):
            doc = message["document"]
            fname = str(doc.get("file_name") or "").lower()
            mime = str(doc.get("mime_type") or "").lower()
            if fname.endswith(".txt") or "text" in mime:
                text = (await self._download_document_text(doc)).strip()
            if not text:
                await self.send_message(chat_id, "❌ Не удалось прочитать текстовый файл.")
                return

        if text.startswith("/start") or text.startswith("/menu") or text.startswith("/mails"):
            await self.show_mailboxes_menu(chat_id, page=0, profile_id=profile_id)
            return

        if text.startswith("/profile") or text.startswith("/id"):
            pstats = self.db.get_profile_stats(profile_id)
            uname_line = f" (@{html.escape(username)})" if username else ""
            txt = (
                f"👤 <b>Ваш изолированный профиль</b>\n"
                f"• Telegram ID: <code>{profile_id}</code>{uname_line}\n"
                f"• Активных почт: <b>{pstats['active']}</b> (токены готовы: <b>{pstats['ready_active']}</b>)\n"
                f"• Использовано: <b>{pstats['used']}</b>\n"
                f"• Удалено: <b>{pstats['deleted']}</b>\n\n"
                f"<i>Каждый Telegram-аккаунт имеет полностью отдельный список почт, "
                f"экспорт, очистку и статистику по своему ID.</i>"
            )
            await self.send_message(chat_id, txt, parse_mode="HTML")
            return

        if text.startswith("/clear"):
            self._cancel_profile_searches(profile_id)
            count = self.db.clear_active_vault_accounts(profile_id)
            await self.send_message(
                chat_id,
                f"🗑 Удалено активных почт в профиле <code>{profile_id}</code>: <b>{count}</b>.",
                parse_mode="HTML",
            )
            await self.show_mailboxes_menu(chat_id, page=0, profile_id=profile_id)
            return

        if text.startswith("/export"):
            await self.export_accounts_txt(chat_id, profile_id=profile_id)
            return

        if text.startswith("/stats"):
            pstats = self.db.get_profile_stats(profile_id)
            cid = self.get_current_client_id()
            txt = (
                f"📊 <b>Статистика профиля</b> <code>{profile_id}</code>:\n"
                f"• Почт в очереди для кодов: <b>{pstats['active']}</b>\n"
                f"• Готовы к выдаче (есть токен): <b>{pstats['ready_active']}/{pstats['active']}</b>\n"
                f"• Использовано почт: <b>{pstats['used']}</b>\n"
                f"• Всего почт в этом профиле: <b>{pstats['total']}</b>\n"
                f"• Протокол: <b>Microsoft Graph API</b>\n"
                f"• Client ID: <code>{html.escape(cid)}</code>"
            )
            await self.send_message(chat_id, txt, parse_mode="HTML")
            return

        if text.startswith("/code"):
            query = text[5:].strip().lower()
            active = self.db.get_active_vault_accounts(profile_id)
            target = next((a for a in active if query and query in a["outlook_email"].lower()), None)
            if target:
                await self.show_mailbox_inbox(chat_id, target["id"], profile_id=profile_id)
            else:
                await self.show_mailboxes_menu(chat_id, page=0, profile_id=profile_id)
            return

        # Otherwise treat message as raw account lines to import into THIS profile
        cleaned_text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
        parsed, skipped = parse_account_lines(cleaned_text)

        if not parsed:
            if "@" in text and len(text.splitlines()) == 1:
                active = self.db.get_active_vault_accounts(profile_id)
                target = next((a for a in active if text.strip().lower() in a["outlook_email"].lower()), None)
                if target:
                    await self.request_code_for_vault(chat_id, target["id"], profile_id=profile_id)
                    return

            await self.send_message(
                chat_id,
                "❌ Не удалось распознать аккаунты.\n\n"
                "Отправьте строки (или .txt файл) вида:\n"
                "<code>tt_user:tt_pass:box@outlook.com:pass:rec@mail.com:rec_pass:TOTP</code>",
                parse_mode="HTML",
            )
            return

        added_count = 0
        for item in parsed:
            self.db.upsert_vault_entry(chat_id=profile_id, **item)
            added_count += 1

        await self.send_message(
            chat_id,
            f"✔ <b>Добавлено {added_count} почт в профиль</b> <code>{profile_id}</code>!",
            parse_mode="HTML",
        )
        await self.show_mailboxes_menu(chat_id, page=0, profile_id=profile_id)

        # Warm up Graph API tokens in background for any brand-new accounts
        asyncio.create_task(self.background_warmup(profile_id, parsed))

    async def run(self) -> None:
        logger.info("Telegram Bot starting...")
        self.http = httpx.AsyncClient(timeout=35.0)
        offset = 0

        # Ensure DB uses Graph API personal App ID (works even when IMAP is disabled)
        self.db.update_settings({"client_id": DEFAULT_APP_ID})

        try:
            me = await self.api_call("getMe")
            bot_user = (me.get("result") or {}).get("username") or "unknown"
            logger.info(f"Connected to Telegram as @{bot_user} (client_id={DEFAULT_APP_ID})")

            # Register convenient Telegram menu commands
            await self.api_call(
                "setMyCommands",
                data={
                    "commands": [
                        {"command": "start", "description": "Список почт"},
                        {"command": "profile", "description": "Мой профиль и ID"},
                        {"command": "clear", "description": "Удалить все почты"},
                        {"command": "export", "description": "Экспорт токенов (.txt)"},
                        {"command": "stats", "description": "Статистика профиля"},
                    ]
                },
            )

            while self.running:
                try:
                    updates = await self.api_call(
                        "getUpdates",
                        data={
                            "offset": offset,
                            "timeout": 20,
                            "allowed_updates": ["message", "callback_query"],
                        },
                    )
                    result = updates.get("result") or []
                    for u in result:
                        offset = max(offset, int(u["update_id"]) + 1)
                        if "message" in u:
                            asyncio.create_task(self.handle_message(u["message"]))
                        elif "callback_query" in u:
                            asyncio.create_task(self.handle_callback_query(u["callback_query"]))
                except asyncio.CancelledError:
                    raise
                except Exception as err:
                    logger.warning(f"Error in Telegram polling: {err}")
                    await asyncio.sleep(2.0)
        finally:
            if self.http:
                await self.http.aclose()


if __name__ == "__main__":
    if not BOT_TOKEN:
        raise SystemExit(
            "TELEGRAM_BOT_TOKEN is not set. Create .env from .env.example or export TELEGRAM_BOT_TOKEN."
        )
    bot = TelegramBot(BOT_TOKEN)
    try:
        asyncio.run(bot.run())
    except KeyboardInterrupt:
        print("\nBot stopped.")
