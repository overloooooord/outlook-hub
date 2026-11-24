"""Async Microsoft OAuth2 (PKCE + refresh token rotation) and Graph API client."""

from __future__ import annotations

import asyncio
import base64
import hashlib
import json
import secrets
import time
from typing import Any
from urllib.parse import urlencode

import httpx

from db import Database
from imap_client import fetch_outlook_imap_messages

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = (
    "offline_access openid profile email "
    "https://graph.microsoft.com/User.Read "
    "https://graph.microsoft.com/Mail.Read"
)
IMAP_SCOPES = "offline_access https://outlook.office.com/IMAP.AccessAsUser.All"
THUNDERBIRD_CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"
THUNDERBIRD_REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient"


class OAuthError(Exception):
    """Raised when Microsoft OAuth2 rejects a token request."""


class GraphAPIError(Exception):
    """Raised when Microsoft Graph returns an error."""

    def __init__(self, status_code: int, message: str) -> None:
        super().__init__(message)
        self.status_code = status_code


def generate_pkce_pair() -> tuple[str, str]:
    """Generate RFC 7636 PKCE (code_verifier, code_challenge)."""
    verifier = secrets.token_urlsafe(64)
    digest = hashlib.sha256(verifier.encode("ascii")).digest()
    challenge = base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")
    return verifier, challenge


def build_authorize_url(
    *,
    client_id: str,
    redirect_uri: str,
    state: str,
    code_challenge: str,
    tenant: str = "common",
    login_hint: str = "",
    prompt: str = "login",
    scope: str = DEFAULT_SCOPES,
) -> str:
    params: dict[str, str] = {
        "client_id": client_id,
        "response_type": "code",
        "redirect_uri": redirect_uri,
        "response_mode": "query",
        "scope": scope,
        "state": state,
        "code_challenge": code_challenge,
        "code_challenge_method": "S256",
    }
    if prompt:
        params["prompt"] = prompt
    if login_hint.strip():
        params["login_hint"] = login_hint.strip()
    return (
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize"
        f"?{urlencode(params)}"
    )


def format_microsoft_error(payload: dict[str, Any], fallback: str = "OAuth error") -> str:
    err = str(payload.get("error") or "")
    desc = str(payload.get("error_description") or "").strip()
    first_line = desc.splitlines()[0] if desc else fallback

    if "AADSTS7000218" in desc:
        return (
            "Приложение в Azure зарегистрировано как Web без client_secret. "
            "В Azure → Authentication добавь платформу «Mobile and desktop applications» "
            "с Redirect URI http://localhost:8765/callback (или укажи client_secret в настройках)."
        )
    if "AADSTS70002" in desc and "mobile" in desc.lower():
        return (
            "Для Device Code Flow в Entra → Authentication → Advanced settings "
            "включи «Allow public client flows = Yes» и нажми Save."
        )
    if "AADSTS50011" in desc:
        return (
            "Redirect URI не совпадает с настройками Azure. "
            "Проверь, что в Azure добавлен точный URI: http://localhost:8765/callback"
        )
    if "AADSTS50194" in desc or "AADSTS9002331" in desc:
        return (
            "Несовпадение типа аккаунтов в Entra. Либо укажи Tenant = consumers в настройках, "
            "либо выбери в Entra тип «Any Entra ID Tenant + Personal Microsoft accounts»."
        )
    if err == "invalid_grant":
        return f"Токен истёк или отозван — требуется повторный вход ({first_line})"
    if err:
        return f"{err}: {first_line}"
    return first_line


def decode_id_token_claims(id_token: str) -> dict[str, Any]:
    """Decode JWT payload without verifying signature (token comes directly from MS token endpoint over TLS)."""
    if not id_token or id_token.count(".") < 2:
        return {}
    try:
        payload_b64 = id_token.split(".")[1]
        padding = "=" * (-len(payload_b64) % 4)
        raw = base64.urlsafe_b64decode(payload_b64 + padding)
        data = json.loads(raw.decode("utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


async def exchange_code_for_tokens(
    http: httpx.AsyncClient,
    *,
    client_id: str,
    code: str,
    redirect_uri: str,
    code_verifier: str,
    tenant: str = "common",
    client_secret: str = "",
    scope: str = DEFAULT_SCOPES,
) -> dict[str, Any]:
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    form: dict[str, str] = {
        "client_id": client_id,
        "scope": scope,
        "code": code,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
        "code_verifier": code_verifier,
    }
    if client_secret:
        form["client_secret"] = client_secret

    resp = await http.post(url, data=form)
    try:
        data = resp.json()
    except Exception:
        data = {}

    if resp.status_code >= 400 or "access_token" not in data:
        raise OAuthError(format_microsoft_error(data, f"HTTP {resp.status_code}"))

    expires_in = int(data.get("expires_in") or 3600)
    return {
        "access_token": str(data["access_token"]),
        "refresh_token": str(data.get("refresh_token") or ""),
        "expires_at": time.time() + expires_in,
        "id_token": str(data.get("id_token") or ""),
    }


async def refresh_access_token(
    http: httpx.AsyncClient,
    *,
    client_id: str,
    refresh_token: str,
    tenant: str = "common",
    client_secret: str = "",
    scope: str = DEFAULT_SCOPES,
) -> dict[str, Any]:
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    form: dict[str, str] = {
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": scope,
    }
    if client_secret:
        form["client_secret"] = client_secret

    resp = await http.post(url, data=form)
    try:
        data = resp.json()
    except Exception:
        data = {}

    if resp.status_code >= 400 or "access_token" not in data:
        raise OAuthError(format_microsoft_error(data, f"HTTP {resp.status_code}"))

    expires_in = int(data.get("expires_in") or 3600)
    return {
        "access_token": str(data["access_token"]),
        "refresh_token": str(data.get("refresh_token") or refresh_token),
        "expires_at": time.time() + expires_in,
        "id_token": str(data.get("id_token") or ""),
    }


async def fetch_profile(
    http: httpx.AsyncClient, access_token: str
) -> tuple[str, str]:
    resp = await http.get(
        f"{GRAPH_BASE}/me",
        params={"$select": "displayName,mail,userPrincipalName"},
        headers={"Authorization": f"Bearer {access_token}"},
    )
    if resp.status_code >= 400:
        raise GraphAPIError(resp.status_code, f"Graph /me failed: HTTP {resp.status_code}")
    data = resp.json()
    email = str(data.get("mail") or data.get("userPrincipalName") or "").strip().lower()
    display_name = str(data.get("displayName") or "").strip()
    return email, display_name


async def fetch_folder_messages(
    http: httpx.AsyncClient,
    access_token: str,
    *,
    folder: str = "inbox",
    top: int = 30,
) -> list[dict[str, Any]]:
    resp = await http.get(
        f"{GRAPH_BASE}/me/mailFolders/{folder}/messages",
        params={
            "$top": str(top),
            "$orderby": "receivedDateTime desc",
            "$select": (
                "id,subject,from,toRecipients,receivedDateTime,"
                "isRead,hasAttachments,bodyPreview,webLink"
            ),
        },
        headers={"Authorization": f"Bearer {access_token}"},
    )
    if resp.status_code >= 400:
        detail = ""
        try:
            detail = str((resp.json().get("error") or {}).get("message") or "")
        except Exception:
            detail = resp.text[:200]
        raise GraphAPIError(resp.status_code, detail or f"HTTP {resp.status_code}")
    data = resp.json()
    return list(data.get("value") or [])


async def fetch_message_body(
    http: httpx.AsyncClient,
    access_token: str,
    graph_id: str,
) -> dict[str, Any]:
    resp = await http.get(
        f"{GRAPH_BASE}/me/messages/{graph_id}",
        params={"$select": "body,toRecipients,isRead,webLink"},
        headers={"Authorization": f"Bearer {access_token}"},
    )
    if resp.status_code >= 400:
        detail = ""
        try:
            detail = str((resp.json().get("error") or {}).get("message") or "")
        except Exception:
            detail = resp.text[:200]
        raise GraphAPIError(resp.status_code, detail or f"HTTP {resp.status_code}")
    data = resp.json()
    body = data.get("body") or {}
    to_list = [
        ((r.get("emailAddress") or {}).get("address") or "")
        for r in (data.get("toRecipients") or [])
    ]
    return {
        "body_content": str(body.get("content") or ""),
        "body_content_type": str(body.get("contentType") or "text").lower(),
        "to_recipients": ", ".join(x for x in to_list if x),
        "is_read": bool(data.get("isRead")),
    }


async def start_device_code_flow(
    http: httpx.AsyncClient,
    *,
    client_id: str,
    tenant: str = "common",
    scope: str = DEFAULT_SCOPES,
) -> dict[str, Any]:
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode"
    resp = await http.post(url, data={"client_id": client_id, "scope": scope})
    try:
        data = resp.json()
    except Exception:
        data = {}
    if resp.status_code >= 400 or "device_code" not in data:
        raise OAuthError(format_microsoft_error(data, f"HTTP {resp.status_code}"))
    return {
        "device_code": str(data["device_code"]),
        "user_code": str(data["user_code"]),
        "verification_uri": str(
            data.get("verification_uri") or "https://microsoft.com/devicelogin"
        ),
        "expires_in": int(data.get("expires_in") or 900),
        "interval": max(2, int(data.get("interval") or 5)),
        "message": str(data.get("message") or ""),
    }


async def poll_device_code_token(
    http: httpx.AsyncClient,
    *,
    client_id: str,
    device_code: str,
    tenant: str = "common",
) -> tuple[str, dict[str, Any] | None]:
    """Poll Microsoft token endpoint once for a device code.

    Returns ("pending", None), ("slow_down", None), or ("completed", token_dict).
    Raises OAuthError on terminal failure (expired/declined/error).
    """
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    resp = await http.post(
        url,
        data={
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id,
            "device_code": device_code,
        },
    )
    try:
        data = resp.json()
    except Exception:
        data = {}

    if resp.status_code == 200 and "access_token" in data:
        expires_in = int(data.get("expires_in") or 3600)
        return (
            "completed",
            {
                "access_token": str(data["access_token"]),
                "refresh_token": str(data.get("refresh_token") or ""),
                "expires_at": time.time() + expires_in,
                "id_token": str(data.get("id_token") or ""),
            },
        )

    err = str(data.get("error") or "")
    if err == "authorization_pending":
        return ("pending", None)
    if err == "slow_down":
        return ("slow_down", None)
    raise OAuthError(format_microsoft_error(data, f"HTTP {resp.status_code}"))


class AccountSyncService:
    """Coordinates token refresh, concurrency locks, and parallel mail sync across accounts (Graph & IMAP XOAUTH2)."""

    def __init__(
        self,
        db: Database,
        http: httpx.AsyncClient,
        *,
        imap_fetcher=None,
    ) -> None:
        self.db = db
        self.http = http
        self._imap_fetcher = imap_fetcher or fetch_outlook_imap_messages
        self._account_locks: dict[int, asyncio.Lock] = {}
        self._sync_semaphore = asyncio.Semaphore(6)

    def _lock_for(self, account_id: int) -> asyncio.Lock:
        lock = self._account_locks.get(account_id)
        if lock is None:
            lock = asyncio.Lock()
            self._account_locks[account_id] = lock
        return lock

    def _resolve_oauth_config(
        self, account: dict[str, Any]
    ) -> tuple[str, str, str, str]:
        settings = self.db.get_settings()
        client_id = (account.get("client_id") or settings["client_id"]).strip()
        client_secret = (
            account.get("client_secret")
            if account.get("client_id")
            else settings["client_secret"]
        ) or ""
        tenant = (account.get("tenant") or settings["tenant"] or "common").strip()
        protocol = (account.get("protocol") or "graph").strip().lower()
        return client_id, client_secret.strip(), tenant, protocol

    async def get_valid_access_token(
        self, account_id: int, *, force_refresh: bool = False
    ) -> str:
        async with self._lock_for(account_id):
            account = self.db.get_account(account_id, include_tokens=True)
            if not account:
                raise OAuthError("Аккаунт не найден")

            now = time.time()
            if (
                not force_refresh
                and account.get("access_token")
                and float(account.get("token_expires_at") or 0) > now + 120
            ):
                return str(account["access_token"])

            client_id, client_secret, tenant, protocol = self._resolve_oauth_config(
                account
            )
            if not client_id:
                raise OAuthError("Не указан Azure Client ID в настройках")

            scope = IMAP_SCOPES if protocol == "imap" else DEFAULT_SCOPES
            token_data = await refresh_access_token(
                self.http,
                client_id=client_id,
                refresh_token=str(account["refresh_token"]),
                tenant=tenant,
                client_secret=client_secret,
                scope=scope,
            )
            self.db.update_account_tokens(
                account_id,
                access_token=token_data["access_token"],
                token_expires_at=token_data["expires_at"],
                refresh_token=token_data["refresh_token"],
            )
            return token_data["access_token"]

    async def sync_account(self, account_id: int) -> dict[str, Any]:
        async with self._sync_semaphore:
            settings = self.db.get_settings()
            top = settings["messages_per_folder"]
            folders = ["inbox", "junkemail"] if settings["sync_junk"] else ["inbox"]
            account = self.db.get_account(account_id, include_tokens=False)
            if not account:
                return {
                    "account_id": account_id,
                    "ok": False,
                    "fetched": 0,
                    "error": "Аккаунт не найден",
                }
            protocol = (account.get("protocol") or "graph").strip().lower()
            email_addr = str(account["email"])

            try:
                token = await self.get_valid_access_token(account_id)
                fetched_total = 0

                for folder in folders:
                    folder_top = top if folder == "inbox" else min(15, top)
                    if protocol == "imap":
                        items = await asyncio.to_thread(
                            self._imap_fetcher,
                            email_addr=email_addr,
                            access_token=token,
                            folder=folder,
                            top=folder_top,
                        )
                    else:
                        try:
                            items = await fetch_folder_messages(
                                self.http, token, folder=folder, top=folder_top
                            )
                        except GraphAPIError as exc:
                            if exc.status_code == 401:
                                token = await self.get_valid_access_token(
                                    account_id, force_refresh=True
                                )
                                items = await fetch_folder_messages(
                                    self.http, token, folder=folder, top=folder_top
                                )
                            elif exc.status_code == 404 and folder == "junkemail":
                                items = []
                            else:
                                raise
                    fetched_total += self.db.upsert_messages(account_id, folder, items)

                self.db.update_account_sync_status(account_id, last_sync_at=time.time())
                return {
                    "account_id": account_id,
                    "ok": True,
                    "fetched": fetched_total,
                    "error": "",
                }
            except Exception as exc:
                err_msg = str(exc) or exc.__class__.__name__
                self.db.update_account_sync_status(account_id, last_error=err_msg)
                return {
                    "account_id": account_id,
                    "ok": False,
                    "fetched": 0,
                    "error": err_msg,
                }

    async def sync_all_accounts(self) -> list[dict[str, Any]]:
        accounts = self.db.list_accounts(include_tokens=False)
        if not accounts:
            return []
        tasks = [self.sync_account(int(acc["id"])) for acc in accounts]
        return list(await asyncio.gather(*tasks))

    async def load_message_body(
        self, message_id: int, *, force_refresh: bool = False
    ) -> dict[str, Any] | None:
        msg = self.db.get_message(message_id)
        if not msg:
            return None

        if msg.get("body_content") is not None and not force_refresh:
            return msg

        account_id = int(msg["account_id"])
        account = self.db.get_account(account_id, include_tokens=False)
        if account and (account.get("protocol") or "graph") == "imap":
            return msg

        token = await self.get_valid_access_token(account_id)
        try:
            body_data = await fetch_message_body(self.http, token, str(msg["graph_id"]))
        except GraphAPIError as exc:
            if exc.status_code == 401:
                token = await self.get_valid_access_token(account_id, force_refresh=True)
                body_data = await fetch_message_body(
                    self.http, token, str(msg["graph_id"])
                )
            else:
                raise

        self.db.update_message_body(
            message_id,
            body_content=body_data["body_content"],
            body_content_type=body_data["body_content_type"],
            to_recipients=body_data["to_recipients"] or None,
            is_read=body_data["is_read"],
        )
        return self.db.get_message(message_id)
