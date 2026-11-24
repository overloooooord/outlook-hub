"""Outlook Hub — local unified web inbox for multiple Outlook accounts via OAuth2 refresh tokens."""

from __future__ import annotations

import asyncio
import base64
import html
import os
import secrets
import time
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

import httpx
import uvicorn
from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from db import Database
from imap_client import RecoveryMailError, fetch_recovery_messages
from outlook_client import (
    DEFAULT_SCOPES,
    IMAP_SCOPES,
    THUNDERBIRD_CLIENT_ID,
    THUNDERBIRD_REDIRECT_URI,
    AccountSyncService,
    GraphAPIError,
    OAuthError,
    build_authorize_url,
    decode_id_token_claims,
    exchange_code_for_tokens,
    fetch_profile,
    format_microsoft_error,
    generate_pkce_pair,
    poll_device_code_token,
    refresh_access_token,
    start_device_code_flow,
)
from vault_utils import generate_totp, parse_account_lines
from auto_login import auto_login_account

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
DEFAULT_HOST = os.environ.get("OUTLOOK_HUB_HOST", "127.0.0.1")
DEFAULT_PORT = int(os.environ.get("OUTLOOK_HUB_PORT", "8765"))
DEFAULT_PUBLIC_URL = os.environ.get("OUTLOOK_HUB_PUBLIC_URL", "").strip()
DEFAULT_PASSWORD = os.environ.get("OUTLOOK_HUB_PASSWORD", "").strip()
DEFAULT_DB_PATH = Path(
    os.environ.get("OUTLOOK_HUB_DB", str(BASE_DIR / "data" / "outlook_hub.db"))
)


class SettingsUpdate(BaseModel):
    client_id: str | None = None
    client_secret: str | None = None
    clear_client_secret: bool = False
    tenant: str | None = None
    poll_interval_sec: int | None = Field(default=None, ge=15, le=3600)
    auto_sync: bool | None = None
    messages_per_folder: int | None = Field(default=None, ge=5, le=100)
    sync_junk: bool | None = None


class OAuthStartRequest(BaseModel):
    login_hint: str = ""
    prompt: str = "login"
    protocol: str = "graph"  # graph | imap


class DeviceCodeStartRequest(BaseModel):
    protocol: str = "graph"  # graph | imap
    client_id: str = ""
    tenant: str = ""


class ManualImportRequest(BaseModel):
    refresh_token: str = Field(min_length=10)
    client_id: str = ""
    client_secret: str = ""
    tenant: str = ""
    email: str = ""
    protocol: str = "auto"  # auto | graph | imap


class BulkTokenImportRequest(BaseModel):
    raw_text: str = Field(min_length=10)
    protocol: str = "auto"  # auto | graph | imap


class VaultImportRequest(BaseModel):
    raw_text: str = Field(min_length=3)
    auto_login: bool = False


def _render_callback_page(*, ok: bool, title: str, detail: str, email: str = "") -> str:
    safe_title = html.escape(title)
    safe_detail = html.escape(detail)
    safe_email = html.escape(email)
    accent = "#22c55e" if ok else "#ef4444"
    icon = "&#10004;" if ok else "&#10006;"
    post_msg = (
        f"window.opener.postMessage({{type: 'outlook-oauth-done', ok: true, email: {safe_email!r}}}, window.location.origin);"
        if ok
        else f"window.opener.postMessage({{type: 'outlook-oauth-done', ok: false, error: {safe_detail!r}}}, window.location.origin);"
    )
    return f"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{safe_title} — Outlook Hub</title>
  <style>
    body {{
      margin: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: #0d1117; color: #e6edf3; min-height: 100vh;
      display: flex; align-items: center; justify-content: center; padding: 24px;
    }}
    .card {{
      max-width: 520px; width: 100%; background: #161b22;
      border: 1px solid #30363d; border-radius: 14px; padding: 28px; text-align: center;
      box-shadow: 0 12px 32px rgba(0, 0, 0, 0.4);
    }}
    .badge {{
      width: 56px; height: 56px; border-radius: 50%; margin: 0 auto 16px;
      display: flex; align-items: center; justify-content: center;
      font-size: 28px; color: #fff; background: {accent};
    }}
    h1 {{ margin: 0 0 10px; font-size: 20px; }}
    p {{ margin: 0 0 18px; color: #8b949e; line-height: 1.5; word-break: break-word; }}
    .email {{ color: #58a6ff; font-weight: 600; }}
    button {{
      background: #238636; color: #fff; border: 0; border-radius: 8px;
      padding: 10px 18px; font-size: 14px; font-weight: 600; cursor: pointer;
    }}
    button:hover {{ background: #2ea043; }}
  </style>
</head>
<body>
  <div class="card">
    <div class="badge">{icon}</div>
    <h1>{safe_title}</h1>
    {f'<p class="email">{safe_email}</p>' if safe_email else ''}
    <p>{safe_detail}</p>
    <button onclick="window.close()">Закрыть вкладку</button>
  </div>
  <script>
    try {{
      if (window.opener && !window.opener.closed) {{
        {post_msg}
      }}
    }} catch (e) {{}}
  </script>
</body>
</html>"""


def create_app(
    *,
    db_path: str | Path | None = None,
    port: int = DEFAULT_PORT,
    http_client: httpx.AsyncClient | None = None,
    enable_background_sync: bool = True,
    recovery_fetcher=None,
    imap_fetcher=None,
    auth_password: str | None = None,
    public_url: str | None = None,
) -> FastAPI:
    db = Database(db_path or DEFAULT_DB_PATH)
    fetcher = recovery_fetcher or fetch_recovery_messages
    effective_public = (
        DEFAULT_PUBLIC_URL if public_url is None else public_url.strip()
    )
    redirect_uri = (
        f"{effective_public.rstrip('/')}/callback"
        if effective_public
        else f"http://localhost:{port}/callback"
    )
    effective_password = (
        DEFAULT_PASSWORD if auth_password is None else auth_password.strip()
    )
    oauth_states: dict[str, dict[str, Any]] = {}
    device_tasks: set[asyncio.Task[Any]] = set()
    sync_lock = asyncio.Lock()
    state_holder: dict[str, Any] = {
        "syncing": False,
        "last_global_sync": 0.0,
        "http": http_client,
        "owns_http": http_client is None,
        "sync_service": None,
        "bg_task": None,
    }

    async def _auto_sync_loop() -> None:
        while True:
            await asyncio.sleep(5)
            try:
                settings = db.get_settings()
                if not settings["auto_sync"]:
                    continue
                interval = settings["poll_interval_sec"]
                if time.time() - state_holder["last_global_sync"] < interval:
                    continue
                if sync_lock.locked():
                    continue
                accounts = db.list_accounts(include_tokens=False)
                if not accounts:
                    state_holder["last_global_sync"] = time.time()
                    continue
                async with sync_lock:
                    state_holder["syncing"] = True
                    try:
                        await state_holder["sync_service"].sync_all_accounts()
                        state_holder["last_global_sync"] = time.time()
                    finally:
                        state_holder["syncing"] = False
            except asyncio.CancelledError:
                raise
            except Exception:
                pass

    @asynccontextmanager
    async def lifespan(_app: FastAPI):
        if state_holder["http"] is None:
            state_holder["http"] = httpx.AsyncClient(timeout=25.0)
        state_holder["sync_service"] = AccountSyncService(
            db, state_holder["http"], imap_fetcher=imap_fetcher
        )
        if enable_background_sync:
            state_holder["bg_task"] = asyncio.create_task(_auto_sync_loop())
        try:
            yield
        finally:
            for t in list(device_tasks):
                t.cancel()
            bg_task = state_holder.get("bg_task")
            if bg_task is not None:
                bg_task.cancel()
                try:
                    await bg_task
                except asyncio.CancelledError:
                    pass
            if state_holder["owns_http"] and state_holder["http"] is not None:
                await state_holder["http"].aclose()

    app = FastAPI(title="Outlook Hub", lifespan=lifespan)

    if effective_password:

        @app.middleware("http")
        async def basic_auth_middleware(request: Request, call_next):
            # Allow Microsoft's OAuth redirect to hit /callback without Basic Auth header
            if request.url.path == "/callback":
                return await call_next(request)
            auth_header = request.headers.get("Authorization", "")
            if auth_header.startswith("Basic "):
                try:
                    decoded = base64.b64decode(auth_header[6:]).decode("utf-8")
                    _, _, provided_pass = decoded.partition(":")
                    if secrets.compare_digest(provided_pass, effective_password):
                        return await call_next(request)
                except Exception:
                    pass
            return Response(
                content="Authentication required",
                status_code=401,
                headers={"WWW-Authenticate": 'Basic realm="Outlook Hub"'},
            )

    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

    def _prune_oauth_states() -> None:
        cutoff = time.time() - 1200
        expired = [k for k, v in oauth_states.items() if v["created_at"] < cutoff]
        for key in expired:
            oauth_states.pop(key, None)

    def _get_service() -> AccountSyncService:
        svc = state_holder.get("sync_service")
        if svc is None:
            if state_holder["http"] is None:
                state_holder["http"] = httpx.AsyncClient(timeout=25.0)
            svc = AccountSyncService(
                db, state_holder["http"], imap_fetcher=imap_fetcher
            )
            state_holder["sync_service"] = svc
        return svc

    @app.get("/", response_class=HTMLResponse)
    async def index() -> HTMLResponse:
        return HTMLResponse((STATIC_DIR / "index.html").read_text(encoding="utf-8"))

    @app.get("/api/settings")
    async def get_settings() -> dict[str, Any]:
        s = db.get_settings()
        return {
            "client_id": s["client_id"],
            "has_client_secret": bool(s["client_secret"]),
            "tenant": s["tenant"],
            "poll_interval_sec": s["poll_interval_sec"],
            "auto_sync": s["auto_sync"],
            "messages_per_folder": s["messages_per_folder"],
            "sync_junk": s["sync_junk"],
            "redirect_uri": redirect_uri,
            "configured": bool(s["client_id"]),
        }

    @app.put("/api/settings")
    async def put_settings(payload: SettingsUpdate) -> dict[str, Any]:
        updates: dict[str, Any] = {
            "client_id": payload.client_id,
            "tenant": payload.tenant,
            "poll_interval_sec": payload.poll_interval_sec,
            "auto_sync": payload.auto_sync,
            "messages_per_folder": payload.messages_per_folder,
            "sync_junk": payload.sync_junk,
        }
        if payload.clear_client_secret:
            updates["client_secret"] = ""
        elif payload.client_secret:
            updates["client_secret"] = payload.client_secret

        db.update_settings(updates)
        return await get_settings()

    @app.post("/api/oauth/start")
    async def oauth_start(payload: OAuthStartRequest) -> dict[str, Any]:
        _prune_oauth_states()
        settings = db.get_settings()
        if not settings["client_id"]:
            raise HTTPException(
                status_code=400,
                detail="Сначала укажи Azure Application (client) ID в настройках.",
            )

        state = secrets.token_urlsafe(24)
        code_verifier, code_challenge = generate_pkce_pair()
        is_thunderbird = (settings["client_id"] == THUNDERBIRD_CLIENT_ID)
        effective_redirect = (
            THUNDERBIRD_REDIRECT_URI if is_thunderbird else redirect_uri
        )
        prompt = payload.prompt if payload.prompt in ("login", "select_account", "") else "login"
        protocol = "imap" if is_thunderbird else (payload.protocol if payload.protocol in ("graph", "imap") else "graph")
        scope = IMAP_SCOPES if protocol == "imap" else DEFAULT_SCOPES

        oauth_states[state] = {
            "status": "pending",
            "created_at": time.time(),
            "code_verifier": code_verifier,
            "client_id": settings["client_id"],
            "client_secret": settings["client_secret"],
            "tenant": settings["tenant"],
            "redirect_uri": effective_redirect,
            "protocol": protocol,
            "scope": scope,
            "email_hint": payload.login_hint.strip(),
            "email": "",
            "error": "",
        }
        authorize_url = build_authorize_url(
            client_id=settings["client_id"],
            redirect_uri=effective_redirect,
            state=state,
            code_challenge=code_challenge,
            tenant=settings["tenant"],
            login_hint=payload.login_hint,
            prompt=prompt,
            scope=scope,
        )
        return {
            "state": state,
            "authorize_url": authorize_url,
            "redirect_uri": effective_redirect,
            "protocol": protocol,
        }

    async def _finalize_oauth_tokens(
        svc: AccountSyncService,
        *,
        tokens: dict[str, Any],
        protocol: str,
        client_id_override: str = "",
        client_secret_override: str = "",
        tenant_override: str = "",
        email_hint: str = "",
    ) -> dict[str, Any]:
        if not tokens.get("refresh_token"):
            raise OAuthError(
                "Microsoft не выдал refresh_token. Убедись, что в Entra добавлено право offline_access."
            )
        email = email_hint.strip().lower()
        display_name = ""
        if protocol == "graph":
            try:
                fetched_email, display_name = await fetch_profile(
                    svc.http, tokens["access_token"]
                )
                if fetched_email:
                    email = fetched_email
            except Exception:
                pass
        if not email:
            claims = decode_id_token_claims(tokens.get("id_token") or "")
            email = str(
                claims.get("email")
                or claims.get("preferred_username")
                or claims.get("upn")
                or ""
            ).strip().lower()
            display_name = display_name or str(claims.get("name") or "").strip()
        if not email:
            raise OAuthError(
                "Не удалось определить email аккаунта автоматически — укажи его вручную."
            )
        return db.upsert_account(
            email=email,
            display_name=display_name,
            refresh_token=tokens["refresh_token"],
            access_token=tokens["access_token"],
            token_expires_at=tokens["expires_at"],
            client_id=client_id_override,
            client_secret=client_secret_override,
            tenant=tenant_override,
            protocol=protocol,
        )

    @app.post("/api/oauth/device/start")
    async def oauth_device_start(payload: DeviceCodeStartRequest) -> dict[str, Any]:
        _prune_oauth_states()
        settings = db.get_settings()
        client_id = (payload.client_id or settings["client_id"]).strip()
        tenant = (payload.tenant or settings["tenant"] or "common").strip()
        protocol = payload.protocol if payload.protocol in ("graph", "imap") else "graph"
        if not client_id:
            raise HTTPException(
                status_code=400,
                detail="Сначала укажи Azure Application (client) ID в настройках.",
            )

        scope = (
            f"openid profile email {IMAP_SCOPES}"
            if protocol == "imap"
            else DEFAULT_SCOPES
        )
        svc = _get_service()
        try:
            dc = await start_device_code_flow(
                svc.http, client_id=client_id, tenant=tenant, scope=scope
            )
        except OAuthError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        state = secrets.token_urlsafe(24)
        oauth_states[state] = {
            "status": "pending",
            "created_at": time.time(),
            "client_id": client_id,
            "tenant": tenant,
            "protocol": protocol,
            "email": "",
            "error": "",
        }

        async def _poll_device() -> None:
            interval = dc["interval"]
            deadline = time.time() + dc["expires_in"]
            try:
                while time.time() < deadline:
                    await asyncio.sleep(interval)
                    status, tokens = await poll_device_code_token(
                        svc.http,
                        client_id=client_id,
                        device_code=dc["device_code"],
                        tenant=tenant,
                    )
                    if status == "slow_down":
                        interval += 5
                        continue
                    if status == "pending":
                        continue
                    if status == "completed" and tokens is not None:
                        account = await _finalize_oauth_tokens(
                            svc,
                            tokens=tokens,
                            protocol=protocol,
                            client_id_override=payload.client_id.strip(),
                            tenant_override=payload.tenant.strip(),
                        )
                        oauth_states[state]["status"] = "completed"
                        oauth_states[state]["email"] = account["email"]
                        asyncio.create_task(svc.sync_account(int(account["id"])))
                        return
                oauth_states[state]["status"] = "error"
                oauth_states[state]["error"] = "Время действия кода устройства истекло."
            except asyncio.CancelledError:
                raise
            except Exception as exc:
                oauth_states[state]["status"] = "error"
                oauth_states[state]["error"] = str(exc) or exc.__class__.__name__

        task = asyncio.create_task(_poll_device())
        device_tasks.add(task)
        task.add_done_callback(device_tasks.discard)

        return {
            "state": state,
            "user_code": dc["user_code"],
            "verification_uri": dc["verification_uri"],
            "expires_in": dc["expires_in"],
            "protocol": protocol,
        }

    @app.get("/api/oauth/status/{state}")
    async def oauth_status(state: str) -> dict[str, Any]:
        entry = oauth_states.get(state)
        if not entry:
            return {"status": "expired", "email": "", "error": ""}
        return {
            "status": entry["status"],
            "email": entry["email"],
            "error": entry["error"],
        }

    @app.get("/callback", response_class=HTMLResponse)
    async def oauth_callback(
        request: Request,
        code: str | None = None,
        state: str | None = None,
        error: str | None = None,
        error_description: str | None = None,
    ) -> HTMLResponse:
        if not state or state not in oauth_states:
            return HTMLResponse(
                _render_callback_page(
                    ok=False,
                    title="Сессия авторизации истекла",
                    detail="Открой Outlook Hub, нажми «Добавить аккаунт» и попробуй снова.",
                ),
                status_code=400,
            )

        entry = oauth_states[state]
        if error:
            msg = format_microsoft_error(
                {"error": error, "error_description": error_description or ""}
            )
            entry["status"] = "error"
            entry["error"] = msg
            return HTMLResponse(
                _render_callback_page(ok=False, title="Вход не выполнен", detail=msg),
                status_code=400,
            )

        if not code:
            msg = "Microsoft не вернул код авторизации."
            entry["status"] = "error"
            entry["error"] = msg
            return HTMLResponse(
                _render_callback_page(ok=False, title="Ошибка авторизации", detail=msg),
                status_code=400,
            )

        svc = _get_service()
        try:
            tokens = await exchange_code_for_tokens(
                svc.http,
                client_id=entry["client_id"],
                code=code,
                redirect_uri=entry["redirect_uri"],
                code_verifier=entry["code_verifier"],
                tenant=entry["tenant"],
                client_secret=entry["client_secret"],
                scope=entry.get("scope") or DEFAULT_SCOPES,
            )
            account = await _finalize_oauth_tokens(
                svc,
                tokens=tokens,
                protocol=entry.get("protocol") or "graph",
                client_id_override=entry["client_id"],
                client_secret_override=entry["client_secret"],
                tenant_override=entry["tenant"],
                email_hint=entry.get("email_hint") or "",
            )
            entry["status"] = "completed"
            entry["email"] = account["email"]
            asyncio.create_task(svc.sync_account(int(account["id"])))

            return HTMLResponse(
                _render_callback_page(
                    ok=True,
                    title="Аккаунт подключён!",
                    email=account["email"],
                    detail="Refresh token сохранён локально, письма уже загружаются. Эту вкладку можно закрыть.",
                )
            )
        except Exception as exc:
            err_msg = str(exc) or exc.__class__.__name__
            entry["status"] = "error"
            entry["error"] = err_msg
            return HTMLResponse(
                _render_callback_page(
                    ok=False, title="Не удалось подключить аккаунт", detail=err_msg
                ),
                status_code=400,
            )

    async def _import_single_token(
        svc: AccountSyncService,
        *,
        refresh_token: str,
        client_id: str,
        client_secret: str,
        tenant: str,
        email: str,
        protocol: str,
        store_custom_client_id: str,
        store_custom_secret: str,
        store_custom_tenant: str,
    ) -> dict[str, Any]:
        chosen_protocol = protocol if protocol in ("graph", "imap") else "graph"
        tokens: dict[str, Any]
        if protocol == "imap":
            tokens = await refresh_access_token(
                svc.http,
                client_id=client_id,
                refresh_token=refresh_token,
                tenant=tenant,
                client_secret=client_secret,
                scope=IMAP_SCOPES,
            )
            chosen_protocol = "imap"
        elif protocol == "graph":
            tokens = await refresh_access_token(
                svc.http,
                client_id=client_id,
                refresh_token=refresh_token,
                tenant=tenant,
                client_secret=client_secret,
                scope=DEFAULT_SCOPES,
            )
            chosen_protocol = "graph"
        else:
            # Auto-detect: try Graph first, fall back to IMAP XOAUTH2 scope
            try:
                tokens = await refresh_access_token(
                    svc.http,
                    client_id=client_id,
                    refresh_token=refresh_token,
                    tenant=tenant,
                    client_secret=client_secret,
                    scope=DEFAULT_SCOPES,
                )
                chosen_protocol = "graph"
            except OAuthError:
                tokens = await refresh_access_token(
                    svc.http,
                    client_id=client_id,
                    refresh_token=refresh_token,
                    tenant=tenant,
                    client_secret=client_secret,
                    scope=IMAP_SCOPES,
                )
                chosen_protocol = "imap"

        account = await _finalize_oauth_tokens(
            svc,
            tokens=tokens,
            protocol=chosen_protocol,
            client_id_override=store_custom_client_id,
            client_secret_override=store_custom_secret,
            tenant_override=store_custom_tenant,
            email_hint=email,
        )
        sync_res = await svc.sync_account(int(account["id"]))
        return {"account": account, "sync": sync_res}

    @app.post("/api/accounts/import")
    async def import_account(payload: ManualImportRequest) -> dict[str, Any]:
        settings = db.get_settings()
        client_id = (payload.client_id or settings["client_id"]).strip()
        client_secret = (
            payload.client_secret.strip()
            if payload.client_id.strip()
            else (payload.client_secret.strip() or settings["client_secret"])
        )
        tenant = (payload.tenant or settings["tenant"] or "common").strip()
        if not client_id:
            raise HTTPException(
                status_code=400,
                detail="Укажи client_id (в форме импорта или в настройках Azure).",
            )
        try:
            return await _import_single_token(
                _get_service(),
                refresh_token=payload.refresh_token.strip(),
                client_id=client_id,
                client_secret=client_secret,
                tenant=tenant,
                email=payload.email.strip(),
                protocol=payload.protocol,
                store_custom_client_id=payload.client_id.strip(),
                store_custom_secret=payload.client_secret.strip(),
                store_custom_tenant=payload.tenant.strip(),
            )
        except HTTPException:
            raise
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/api/accounts/import-bulk")
    async def import_bulk_tokens(payload: BulkTokenImportRequest) -> dict[str, Any]:
        import re

        uuid_re = re.compile(
            r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"
        )
        settings = db.get_settings()
        svc = _get_service()
        imported: list[dict[str, Any]] = []
        errors: list[dict[str, str]] = []

        for raw_line in payload.raw_text.splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            parts = [p.strip() for p in line.split(":") if p.strip()]
            email = next((p for p in parts if "@" in p and "." in p), "")
            cid = next((p for p in parts if uuid_re.match(p)), "")
            # Refresh token is usually the longest token (e.g. starts with M. or 0. or len >= 30)
            candidates = [p for p in parts if p != email and p != cid]
            rt = max(candidates, key=len) if candidates else ""
            if len(rt) < 15:
                errors.append({"line": line[:60], "error": "Не найден refresh_token"})
                continue
            effective_cid = cid or settings["client_id"]
            if not effective_cid:
                errors.append(
                    {"line": email or line[:40], "error": "Не указан client_id"}
                )
                continue
            try:
                res = await _import_single_token(
                    svc,
                    refresh_token=rt,
                    client_id=effective_cid,
                    client_secret="" if cid else settings["client_secret"],
                    tenant=settings["tenant"] or "common",
                    email=email,
                    protocol=payload.protocol,
                    store_custom_client_id=cid,
                    store_custom_secret="",
                    store_custom_tenant="",
                )
                imported.append(res)
            except Exception as exc:
                errors.append(
                    {"line": email or line[:40], "error": str(exc)}
                )

        return {
            "imported_count": len(imported),
            "error_count": len(errors),
            "imported": imported,
            "errors": errors,
            "stats": db.get_stats(),
        }

    @app.get("/api/accounts")
    async def list_accounts() -> dict[str, Any]:
        return {
            "accounts": db.list_accounts(include_tokens=False),
            "stats": db.get_stats(),
            "syncing": state_holder["syncing"],
            "last_global_sync": state_holder["last_global_sync"],
        }

    @app.post("/api/accounts/{account_id}/sync")
    async def sync_single_account(account_id: int) -> dict[str, Any]:
        if not db.get_account(account_id):
            raise HTTPException(status_code=404, detail="Аккаунт не найден")
        result = await _get_service().sync_account(account_id)
        return {
            "result": result,
            "accounts": db.list_accounts(include_tokens=False),
            "stats": db.get_stats(),
        }

    @app.delete("/api/accounts/{account_id}")
    async def delete_account(account_id: int) -> dict[str, Any]:
        deleted = db.delete_account(account_id)
        if not deleted:
            raise HTTPException(status_code=404, detail="Аккаунт не найден")
        return {"ok": True, "stats": db.get_stats()}

    @app.post("/api/sync")
    async def sync_all() -> dict[str, Any]:
        if sync_lock.locked():
            return {
                "results": [],
                "already_running": True,
                "accounts": db.list_accounts(include_tokens=False),
                "stats": db.get_stats(),
            }
        async with sync_lock:
            state_holder["syncing"] = True
            try:
                results = await _get_service().sync_all_accounts()
                state_holder["last_global_sync"] = time.time()
            finally:
                state_holder["syncing"] = False
        return {
            "results": results,
            "already_running": False,
            "accounts": db.list_accounts(include_tokens=False),
            "stats": db.get_stats(),
            "last_global_sync": state_holder["last_global_sync"],
        }

    @app.get("/api/messages")
    async def list_messages(
        account_id: int | None = Query(default=None),
        folder: str = Query(default="all"),
        unread_only: bool = Query(default=False),
        search: str = Query(default=""),
        limit: int = Query(default=150, ge=1, le=500),
        offset: int = Query(default=0, ge=0),
    ) -> dict[str, Any]:
        messages = db.list_messages(
            account_id=account_id,
            folder=folder,
            unread_only=unread_only,
            search=search,
            limit=limit,
            offset=offset,
        )
        return {"messages": messages, "stats": db.get_stats()}

    @app.get("/api/messages/{message_id}")
    async def get_message(
        message_id: int, refresh: bool = Query(default=False)
    ) -> dict[str, Any]:
        try:
            msg = await _get_service().load_message_body(
                message_id, force_refresh=refresh
            )
        except (OAuthError, GraphAPIError) as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc
        if not msg:
            raise HTTPException(status_code=404, detail="Письмо не найдено")
        return {"message": msg}

    def _attach_totp(entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
        now = time.time()
        for entry in entries:
            code, remaining = generate_totp(entry.get("totp_secret") or "", timestamp=now)
            entry["totp_code"] = code
            entry["totp_remaining"] = remaining
        return entries

    @app.get("/api/vault")
    async def list_vault() -> dict[str, Any]:
        entries = _attach_totp(db.list_vault_entries())
        return {"entries": entries, "stats": db.get_stats()}

    @app.post("/api/vault/import")
    async def import_vault(payload: VaultImportRequest) -> dict[str, Any]:
        parsed, skipped = parse_account_lines(payload.raw_text)
        if not parsed:
            raise HTTPException(
                status_code=400,
                detail="Не удалось распознать ни одной строки с email. Проверь формат.",
            )
        imported_ids = []
        for item in parsed:
            row = db.upsert_vault_entry(**item)
            imported_ids.append(int(row["id"]))

        entries = _attach_totp(db.list_vault_entries())

        # If user asked for immediate auto-login, launch background task
        if payload.auto_login:
            async def _bg_auto_login_batch(e_ids: list[int]):
                svc = _get_service()
                for eid in e_ids:
                    entry = db.get_vault_entry(eid)
                    if not entry:
                        continue
                    try:
                        await asyncio.to_thread(
                            auto_login_account,
                            entry,
                            base_url=f"http://127.0.0.1:{port}",
                            headless=True,
                        )
                        connected = db.get_account_by_email(entry["outlook_email"])
                        if connected:
                            asyncio.create_task(svc.sync_account(int(connected["id"])))
                    except Exception as err:
                        logging.getLogger("outlook_hub").warning(
                            f"Background auto-login failed for {entry.get('outlook_email')}: {err}"
                        )
            asyncio.create_task(_bg_auto_login_batch(imported_ids))

        return {
            "imported_count": len(parsed),
            "skipped": skipped,
            "entries": entries,
            "auto_login_started": payload.auto_login,
            "stats": db.get_stats(),
        }

    @app.post("/api/vault/{entry_id}/auto-login")
    async def auto_login_vault_entry(entry_id: int) -> dict[str, Any]:
        entry = db.get_vault_entry(entry_id)
        if not entry:
            raise HTTPException(status_code=404, detail="Запись не найдена в хранилище.")
        try:
            await asyncio.to_thread(
                auto_login_account,
                entry,
                base_url=f"http://127.0.0.1:{port}",
                headless=True,
            )
            connected = db.get_account_by_email(entry["outlook_email"])
            if connected:
                asyncio.create_task(_get_service().sync_account(int(connected["id"])))
            return {
                "ok": True,
                "email": entry["outlook_email"],
                "account": connected,
                "stats": db.get_stats(),
            }
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/api/vault/auto-login-all")
    async def auto_login_vault_all(force: bool = Query(default=False)) -> dict[str, Any]:
        entries = db.list_vault_entries()
        targets = [e for e in entries if force or not e["oauth_connected"]]
        if not targets:
            return {"ok": True, "total": 0, "processed": 0, "results": []}

        svc = _get_service()
        results = []
        for e in targets:
            try:
                await asyncio.to_thread(
                    auto_login_account,
                    e,
                    base_url=f"http://127.0.0.1:{port}",
                    headless=True,
                )
                connected = db.get_account_by_email(e["outlook_email"])
                if connected:
                    asyncio.create_task(svc.sync_account(int(connected["id"])))
                results.append({"email": e["outlook_email"], "status": "ok", "error": ""})
            except Exception as exc:
                results.append({"email": e["outlook_email"], "status": "error", "error": str(exc)})

        return {
            "ok": True,
            "total": len(targets),
            "processed": len(results),
            "results": results,
            "stats": db.get_stats(),
        }

    @app.post("/api/vault/{entry_id}/recovery-mail")
    async def fetch_vault_recovery_mail(
        entry_id: int,
        imap_host: str = Query(default=""),
        limit: int = Query(default=8, ge=1, le=20),
    ) -> dict[str, Any]:
        entry = db.get_vault_entry(entry_id)
        if not entry:
            raise HTTPException(status_code=404, detail="Запись не найдена")
        rec_email = (entry.get("recovery_email") or "").strip()
        rec_pass = (entry.get("recovery_password") or "").strip()
        if not rec_email or not rec_pass:
            raise HTTPException(
                status_code=400,
                detail="У этой записи не указана резервная почта или пароль.",
            )
        try:
            result = await asyncio.to_thread(
                fetcher,
                email_addr=rec_email,
                password=rec_pass,
                imap_host=imap_host,
                limit=limit,
            )
            return result
        except RecoveryMailError as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc

    @app.delete("/api/vault/{entry_id}")
    async def delete_vault(entry_id: int) -> dict[str, Any]:
        deleted = db.delete_vault_entry(entry_id)
        if not deleted:
            raise HTTPException(status_code=404, detail="Запись не найдена")
        return {"ok": True, "stats": db.get_stats()}

    @app.get("/api/export-tokens")
    async def export_tokens() -> JSONResponse:
        settings = db.get_settings()
        accounts = db.list_accounts(include_tokens=True)
        exported = [
            {
                "email": acc["email"],
                "protocol": acc.get("protocol") or "graph",
                "client_id": acc["client_id"] or settings["client_id"],
                "tenant": acc["tenant"] or settings["tenant"],
                "refresh_token": acc["refresh_token"],
            }
            for acc in accounts
        ]
        return JSONResponse(
            content=exported,
            headers={
                "Content-Disposition": 'attachment; filename="outlook-refresh-tokens.json"'
            },
        )

    return app


app = create_app()


if __name__ == "__main__":
    url_hint = DEFAULT_PUBLIC_URL or f"http://localhost:{DEFAULT_PORT}"
    print(f"Outlook Hub запущен: {url_hint} (bind {DEFAULT_HOST}:{DEFAULT_PORT})")
    uvicorn.run(app, host=DEFAULT_HOST, port=DEFAULT_PORT, log_level="info")
