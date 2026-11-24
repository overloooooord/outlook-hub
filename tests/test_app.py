"""End-to-end unit tests for Outlook Hub (DB, PKCE, OAuth2 callback, token rotation, Graph sync)."""

from __future__ import annotations

import base64
import hashlib
import json
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import httpx
import pytest
from fastapi.testclient import TestClient

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app import create_app
from outlook_client import format_microsoft_error, generate_pkce_pair


def _make_mock_transport():
    state = {
        "token_calls": [],
        "body_calls": 0,
        "force_401_once": False,
        "rotated_refresh_count": 0,
    }

    def handler(request: httpx.Request) -> httpx.Response:
        url = str(request.url)

        if "/oauth2/v2.0/token" in url:
            body = parse_qs(request.content.decode("utf-8"))
            grant_type = (body.get("grant_type") or [""])[0]
            state["token_calls"].append({"grant_type": grant_type, "body": body})

            if grant_type == "authorization_code":
                code = (body.get("code") or [""])[0]
                if code == "bad_code":
                    return httpx.Response(
                        400,
                        json={
                            "error": "invalid_grant",
                            "error_description": "AADSTS7000218: The request body must contain client_secret",
                        },
                    )
                return httpx.Response(
                    200,
                    json={
                        "access_token": f"access_init_{code}",
                        "refresh_token": f"refresh_init_{code}",
                        "expires_in": 3600,
                    },
                )

            if grant_type == "refresh_token":
                rt = (body.get("refresh_token") or [""])[0]
                if rt == "expired_rt":
                    return httpx.Response(
                        400,
                        json={
                            "error": "invalid_grant",
                            "error_description": "AADSTS70000: Refresh token expired",
                        },
                    )
                state["rotated_refresh_count"] += 1
                n = state["rotated_refresh_count"]
                suffix = "acc2" if "acc2" in rt else "acc1"
                return httpx.Response(
                    200,
                    json={
                        "access_token": f"access_refreshed_{suffix}_{n}",
                        "refresh_token": f"refresh_rotated_{suffix}_{n}",
                        "expires_in": 3600,
                    },
                )

        if url.startswith("https://graph.microsoft.com/v1.0/me?"):
            auth = request.headers.get("Authorization", "")
            if "acc2" in auth:
                return httpx.Response(
                    200,
                    json={"displayName": "Second Box", "mail": "second@outlook.com"},
                )
            return httpx.Response(
                200,
                json={"displayName": "First Box", "mail": "first@outlook.com"},
            )

        if "/mailFolders/inbox/messages" in url:
            if state["force_401_once"]:
                state["force_401_once"] = False
                return httpx.Response(
                    401, json={"error": {"message": "Token expired"}}
                )
            auth = request.headers.get("Authorization", "")
            tag = "second" if "acc2" in auth else "first"
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": f"msg_{tag}_1",
                            "subject": f"Код подтверждения для {tag}",
                            "from": {
                                "emailAddress": {
                                    "name": "Service",
                                    "address": "noreply@service.test",
                                }
                            },
                            "toRecipients": [
                                {"emailAddress": {"address": f"{tag}@outlook.com"}}
                            ],
                            "receivedDateTime": "2026-11-22T00:30:00Z",
                            "isRead": False,
                            "hasAttachments": False,
                            "bodyPreview": "Ваш код: 482910",
                            "webLink": f"https://outlook.live.com/mail/{tag}_1",
                        }
                    ]
                },
            )

        if "/mailFolders/junkemail/messages" in url:
            auth = request.headers.get("Authorization", "")
            tag = "second" if "acc2" in auth else "first"
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": f"junk_{tag}_1",
                            "subject": "Письмо в спаме",
                            "from": {
                                "emailAddress": {
                                    "name": "Promo",
                                    "address": "promo@example.com",
                                }
                            },
                            "receivedDateTime": "2026-11-21T22:00:00Z",
                            "isRead": True,
                            "hasAttachments": False,
                            "bodyPreview": "Реклама",
                            "webLink": "",
                        }
                    ]
                },
            )

        if "/v1.0/me/messages/" in url:
            state["body_calls"] += 1
            return httpx.Response(
                200,
                json={
                    "body": {
                        "contentType": "html",
                        "content": "<div>Привет! Ваш код: <b>482910</b></div>",
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": "first@outlook.com"}}
                    ],
                    "isRead": False,
                },
            )

        return httpx.Response(404, json={"error": {"message": f"Unhandled {url}"}})

    return httpx.MockTransport(handler), state


def test_pkce_generation():
    verifier, challenge = generate_pkce_pair()
    assert 43 <= len(verifier) <= 128
    expected = (
        base64.urlsafe_b64encode(hashlib.sha256(verifier.encode("ascii")).digest())
        .rstrip(b"=")
        .decode("ascii")
    )
    assert challenge == expected
    assert "=" not in challenge


def test_error_formatting():
    msg = format_microsoft_error(
        {
            "error": "invalid_client",
            "error_description": "AADSTS7000218: client_secret required",
        }
    )
    assert "Mobile and desktop applications" in msg


def test_full_oauth_sync_reader_and_rotation(tmp_path: Path):
    transport, mock_state = _make_mock_transport()
    http_client = httpx.AsyncClient(transport=transport)
    db_file = tmp_path / "test_outlook.db"
    app = create_app(
        db_path=db_file,
        port=8765,
        http_client=http_client,
        enable_background_sync=False,
    )

    with TestClient(app) as client:
        # 1. Initially unconfigured
        r = client.get("/api/settings")
        assert r.status_code == 200
        assert r.json()["configured"] is False
        assert r.json()["redirect_uri"] == "http://localhost:8765/callback"

        # 2. Save Azure client_id
        r = client.put("/api/settings", json={"client_id": "test-client-id-123"})
        assert r.status_code == 200
        assert r.json()["configured"] is True

        # 3. Start OAuth flow for account #1
        r = client.post(
            "/api/oauth/start",
            json={"login_hint": "first@outlook.com", "prompt": "login"},
        )
        assert r.status_code == 200
        start_data = r.json()
        state_token = start_data["state"]
        auth_url = urlparse(start_data["authorize_url"])
        qs = parse_qs(auth_url.query)
        assert qs["client_id"] == ["test-client-id-123"]
        assert qs["code_challenge_method"] == ["S256"]
        assert qs["login_hint"] == ["first@outlook.com"]
        assert "offline_access" in qs["scope"][0]

        # 4. Status is pending before callback
        assert client.get(f"/api/oauth/status/{state_token}").json()["status"] == "pending"

        # 5. Simulate Microsoft callback
        cb = client.get(f"/callback?code=acc1&state={state_token}")
        assert cb.status_code == 200
        assert "first@outlook.com" in cb.text

        # 6. Status now completed (used by Incognito polling)
        st = client.get(f"/api/oauth/status/{state_token}").json()
        assert st["status"] == "completed"
        assert st["email"] == "first@outlook.com"

        # 7. Manual import for account #2
        imp = client.post(
            "/api/accounts/import",
            json={"refresh_token": "manual_rt_token_acc2", "email": "second@outlook.com"},
        )
        assert imp.status_code == 200
        assert imp.json()["sync"]["ok"] is True

        # 8. Sync all accounts (trigger 401 once to prove automatic refresh token rotation!)
        mock_state["force_401_once"] = True
        sync_res = client.post("/api/sync")
        assert sync_res.status_code == 200
        results = sync_res.json()["results"]
        assert len(results) == 2
        assert all(item["ok"] for item in results)
        assert mock_state["rotated_refresh_count"] >= 2

        # 9. Verify exported tokens contain rotated refresh token
        exported = client.get("/api/export-tokens").json()
        assert len(exported) == 2
        assert any(e["refresh_token"].startswith("refresh_rotated_") for e in exported)

        # 10. Verify unified inbox & junk & search
        inbox = client.get("/api/messages?folder=inbox").json()["messages"]
        assert len(inbox) == 2
        junk = client.get("/api/messages?folder=junkemail").json()["messages"]
        assert len(junk) == 2
        searched = client.get("/api/messages?search=482910").json()["messages"]
        assert len(searched) == 2

        # 11. Load message body once from Graph, second time from SQLite cache
        msg_id = inbox[0]["id"]
        m1 = client.get(f"/api/messages/{msg_id}").json()["message"]
        assert "482910" in m1["body_content"]
        assert m1["body_content_type"] == "html"
        assert mock_state["body_calls"] == 1

        m2 = client.get(f"/api/messages/{msg_id}").json()["message"]
        assert m2["body_content"] == m1["body_content"]
        assert mock_state["body_calls"] == 1  # served from cache!

        # 12. Delete account cascades its messages
        acc_id = client.get("/api/accounts").json()["accounts"][0]["id"]
        assert client.delete(f"/api/accounts/{acc_id}").status_code == 200
        assert len(client.get("/api/accounts").json()["accounts"]) == 1


def test_totp_and_vault_parser(tmp_path: Path):
    from vault_utils import generate_totp, parse_account_lines

    # RFC 6238 SHA1 test vector: secret "12345678901234567890" -> Base32 GEZDGNBVGY3TQOJQGEZDGNBVGY3TQOJQ, t=59 -> 287082
    code, remaining = generate_totp("GEZDGNBVGY3TQOJQGEZDGNBVGY3TQOJQ", timestamp=59)
    assert code == "287082"
    assert remaining == 1

    # Lowercase Base32 secret should work too
    code_low, _ = generate_totp("gezdgnbvgy3tqojq", timestamp=59)
    assert len(code_low) == 6 and code_low.isdigit()

    sample = """
[11/22/26 12:46 AM] testuser:
Forwarded from sender
user111:PassOne!:alpha@outlook.com:MailPass1:rec1@example.com:RecPass1:JBSWY3DPEHPK3PXP
user222:PassTwo!:beta@outlook.com:MailPass2::rec2@example.com:RecPass2:jbswy3dpehpk3pxp
Формат лог:пасс от тт:лог:паспочта::резервная почта:пароль:2фа
"""
    parsed, skipped = parse_account_lines(sample)
    assert len(parsed) == 2
    assert not skipped
    assert parsed[0]["tt_login"] == "user111"
    assert parsed[0]["outlook_email"] == "alpha@outlook.com"
    assert parsed[0]["recovery_email"] == "rec1@example.com"
    assert parsed[1]["outlook_email"] == "beta@outlook.com"
    assert parsed[1]["totp_secret"] == "jbswy3dpehpk3pxp"

    app = create_app(db_path=tmp_path / "vault.db", enable_background_sync=False)
    with TestClient(app) as client:
        res = client.post("/api/vault/import", json={"raw_text": sample})
        assert res.status_code == 200
        data = res.json()
        assert data["imported_count"] == 2
        assert len(data["entries"]) == 2
        assert len(data["entries"][0]["totp_code"]) == 6
        assert data["stats"]["total_vault"] == 2

        entry_id = data["entries"][0]["id"]
        assert client.delete(f"/api/vault/{entry_id}").status_code == 200
        assert len(client.get("/api/vault").json()["entries"]) == 1


def test_recovery_imap_code_extraction_and_endpoint(tmp_path: Path):
    from imap_client import extract_verification_codes, resolve_imap_host

    assert resolve_imap_host("user@legenmail.com") == "imap.firstmail.ltd"
    assert resolve_imap_host("user@rambler.ru") == "imap.rambler.ru"
    assert resolve_imap_host("user@legenmail.com", "custom.imap.local") == "custom.imap.local"

    codes_ms = extract_verification_codes(
        "Microsoft account security code",
        "Please use the following security code for the Microsoft account: 492817\n"
        "Visit https://account.live.com/12345678 for details.",
    )
    assert codes_ms[0] == "492817"

    codes_ru = extract_verification_codes(
        "Код безопасности",
        "Ваш код подтверждения: 7391042. Никому не сообщайте.",
    )
    assert codes_ru[0] == "7391042"

    def fake_fetcher(*, email_addr, password, imap_host="", limit=8):
        return {
            "recovery_email": email_addr,
            "imap_host": imap_host or "imap.firstmail.ltd",
            "messages": [
                {
                    "folder": "INBOX",
                    "subject": "Microsoft account security code",
                    "sender": "account-security-noreply@accountprotection.microsoft.com",
                    "date": "Sun, 22 Nov 2026 01:31:00 +0300",
                    "preview": "Security code: 615243",
                    "text": "Security code: 615243",
                    "codes": ["615243"],
                }
            ],
        }

    app = create_app(
        db_path=tmp_path / "rec.db",
        enable_background_sync=False,
        recovery_fetcher=fake_fetcher,
    )
    with TestClient(app) as client:
        imp = client.post(
            "/api/vault/import",
            json={
                "raw_text": "u1:p1:box@outlook.com:mp1:rec@legenmail.com:rp1:JBSWY3DPEHPK3PXP"
            },
        )
        entry_id = imp.json()["entries"][0]["id"]
        r = client.post(f"/api/vault/{entry_id}/recovery-mail")
        assert r.status_code == 200
        body = r.json()
        assert body["recovery_email"] == "rec@legenmail.com"
        assert body["messages"][0]["codes"] == ["615243"]


def test_device_code_imap_bulk_import_and_vps_auth(tmp_path: Path):
    import time

    poll_hits = {"n": 0}

    def handler(request: httpx.Request) -> httpx.Response:
        url = str(request.url)
        body = parse_qs(request.content.decode("utf-8"))

        if "/oauth2/v2.0/devicecode" in url:
            return httpx.Response(
                200,
                json={
                    "device_code": "dev_code_1",
                    "user_code": "ABCD-EFGH",
                    "verification_uri": "https://microsoft.com/devicelogin",
                    "expires_in": 900,
                    "interval": 1,
                },
            )

        if "/oauth2/v2.0/token" in url:
            grant_type = (body.get("grant_type") or [""])[0]
            if grant_type == "urn:ietf:params:oauth:grant-type:device_code":
                poll_hits["n"] += 1
                return httpx.Response(
                    200,
                    json={
                        "access_token": "dev_access_tok",
                        "refresh_token": "dev_refresh_tok_1234567890",
                        "expires_in": 3600,
                    },
                )
            if grant_type == "refresh_token":
                scope = (body.get("scope") or [""])[0]
                if "graph.microsoft.com" in scope:
                    # Simulate an IMAP-only refresh token rejecting Graph scope
                    return httpx.Response(
                        400,
                        json={
                            "error": "invalid_scope",
                            "error_description": "IMAP-only token",
                        },
                    )
                return httpx.Response(
                    200,
                    json={
                        "access_token": "imap_access_tok",
                        "refresh_token": "imap_refresh_rotated_1234567890",
                        "expires_in": 3600,
                    },
                )

        if url.startswith("https://graph.microsoft.com/v1.0/me"):
            return httpx.Response(
                200,
                json={"displayName": "Device User", "mail": "device@outlook.com"},
            )
        if "/mailFolders/" in url:
            return httpx.Response(200, json={"value": []})

        return httpx.Response(404)

    def fake_imap_fetcher(*, email_addr, access_token, folder="inbox", top=30):
        if folder != "inbox":
            return []
        return [
            {
                "id": "imap:inbox:42",
                "subject": "TikTok verification code",
                "from": {"emailAddress": {"name": "TikTok", "address": "noreply@tiktok.com"}},
                "toRecipients": [{"emailAddress": {"address": email_addr}}],
                "receivedDateTime": "2026-11-23T01:00:00Z",
                "isRead": False,
                "hasAttachments": False,
                "bodyPreview": "Your TikTok code is 777123",
                "bodyContent": "Your TikTok code is 777123",
                "bodyContentType": "text",
                "webLink": "",
            }
        ]

    http_client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    app = create_app(
        db_path=tmp_path / "vps.db",
        http_client=http_client,
        enable_background_sync=False,
        imap_fetcher=fake_imap_fetcher,
        auth_password="my-vps-password",
        public_url="https://mail.example.com",
    )

    with TestClient(app) as client:
        # 1. VPS Basic Auth blocks unauthenticated access, allows /callback
        assert client.get("/api/settings").status_code == 401
        assert client.get("/callback").status_code == 400  # reached handler without 401

        auth_hdr = {
            "Authorization": "Basic "
            + base64.b64encode(b"admin:my-vps-password").decode("ascii")
        }
        s = client.get("/api/settings", headers=auth_hdr)
        assert s.status_code == 200
        assert s.json()["redirect_uri"] == "https://mail.example.com/callback"

        client.put(
            "/api/settings",
            json={"client_id": "11111111-2222-3333-4444-555555555555"},
            headers=auth_hdr,
        )

        # 2. Device Code Flow start + poll completion
        dc = client.post(
            "/api/oauth/device/start",
            json={"protocol": "graph"},
            headers=auth_hdr,
        )
        assert dc.status_code == 200
        dc_data = dc.json()
        assert dc_data["user_code"] == "ABCD-EFGH"

        # Wait briefly for background poller (interval clamped to 2s)
        status = "pending"
        for _ in range(25):
            st = client.get(
                f"/api/oauth/status/{dc_data['state']}", headers=auth_hdr
            ).json()
            status = st["status"]
            if status == "completed":
                assert st["email"] == "device@outlook.com"
                break
            time.sleep(0.15)
        assert status == "completed"

        # 3. Bulk IMAP-only refresh token import with auto-fallback Graph -> IMAP XOAUTH2
        bulk = client.post(
            "/api/accounts/import-bulk",
            json={
                "raw_text": (
                    "imapuser@outlook.com:pass123:"
                    "M.C512_IMAP_ONLY_REFRESH_TOKEN_LONG:"
                    "9e5f94bc-e8a4-4e73-b8be-63364c29d753"
                )
            },
            headers=auth_hdr,
        )
        assert bulk.status_code == 200
        bdata = bulk.json()
        assert bdata["imported_count"] == 1
        assert bdata["imported"][0]["account"]["protocol"] == "imap"
        assert bdata["imported"][0]["sync"]["ok"] is True

        msgs = client.get("/api/messages?search=777123", headers=auth_hdr).json()["messages"]
        assert len(msgs) == 1
        assert msgs[0]["account_email"] == "imapuser@outlook.com"


def test_vault_auto_login_endpoints(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    """Verify single and batch auto-login endpoints with mocked playwright engine."""
    import app as app_module
    from db import Database

    db_path = tmp_path / "test_autologin.db"
    db = Database(db_path)
    # Seed vault account
    v = db.upsert_vault_entry(
        outlook_email="auto_test@outlook.com",
        outlook_password="pass",
        recovery_email="rec@legenmail.com",
        recovery_password="rpass",
        totp_secret="JBSWY3DPEHPK3PXP",
    )
    v_id = v["id"]

    calls: list[str] = []

    def mock_auto_login(vault_entry, **kwargs):
        calls.append(vault_entry["outlook_email"])
        # Mock what the callback does in real flow:
        db.upsert_account(
            email=vault_entry["outlook_email"],
            display_name="Auto User",
            refresh_token="mock_rt_123",
            access_token="mock_at_123",
            token_expires_at=9999999999,
        )
        return {"ok": True, "email": vault_entry["outlook_email"]}

    monkeypatch.setattr(app_module, "auto_login_account", mock_auto_login)

    test_app = create_app(db_path=db_path, enable_background_sync=False)
    with TestClient(test_app) as client:
        # 1. Single auto-login
        resp = client.post(f"/api/vault/{v_id}/auto-login")
        assert resp.status_code == 200
        data = resp.json()
        assert data["ok"] is True
        assert data["email"] == "auto_test@outlook.com"
        assert "auto_test@outlook.com" in calls

        # Verify DB account was created
        acc = db.get_account_by_email("auto_test@outlook.com")
        assert acc is not None

        # 2. Batch auto-login for all
        db.upsert_vault_entry(
            outlook_email="auto_test2@outlook.com",
            outlook_password="pass2",
        )
        resp_all = client.post("/api/vault/auto-login-all")
        assert resp_all.status_code == 200
        data_all = resp_all.json()
        assert data_all["ok"] is True
        assert data_all["total"] == 1  # Only unconn
        assert "auto_test2@outlook.com" in calls


def test_multi_profile_isolation_and_migration(tmp_path: Path):
    """Verify legacy UNIQUE(outlook_email) migration and strict per-ID profile isolation."""
    import asyncio
    import sqlite3
    import time

    from db import Database
    from telegram_bot import TelegramBot

    db_path = tmp_path / "legacy_multi_profile.db"

    # 1. Build a legacy DB with global UNIQUE(outlook_email) and pre-existing rows
    conn = sqlite3.connect(str(db_path))
    conn.executescript(
        """
        CREATE TABLE vault_accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            outlook_email TEXT NOT NULL UNIQUE COLLATE NOCASE,
            outlook_password TEXT NOT NULL DEFAULT '',
            tt_login TEXT NOT NULL DEFAULT '',
            tt_password TEXT NOT NULL DEFAULT '',
            recovery_email TEXT NOT NULL DEFAULT '',
            recovery_password TEXT NOT NULL DEFAULT '',
            totp_secret TEXT NOT NULL DEFAULT '',
            notes TEXT NOT NULL DEFAULT '',
            created_at REAL NOT NULL,
            status TEXT NOT NULL DEFAULT 'active',
            chat_id INTEGER NOT NULL DEFAULT 0
        );
        """
    )
    now = time.time()
    conn.execute(
        "INSERT INTO vault_accounts (id, outlook_email, outlook_password, created_at, status, chat_id) "
        "VALUES (10, 'legacy_web@outlook.com', 'p0', ?, 'active', 0)",
        (now,),
    )
    conn.execute(
        "INSERT INTO vault_accounts (id, outlook_email, outlook_password, created_at, status, chat_id) "
        "VALUES (20, 'p1_only@outlook.com', 'p1', ?, 'active', 111111)",
        (now,),
    )
    conn.execute(
        "INSERT INTO vault_accounts (id, outlook_email, outlook_password, created_at, status, chat_id) "
        "VALUES (30, 'p2_only@outlook.com', 'p2', ?, 'active', 222222)",
        (now,),
    )
    conn.commit()
    conn.close()

    # 2. Opening Database triggers safe atomic migration preserving all rows & IDs
    db = Database(db_path)
    p1_rows = db.get_active_vault_accounts(111111)
    p2_rows = db.get_active_vault_accounts(222222)
    assert [r["id"] for r in p1_rows] == [20]
    assert [r["id"] for r in p2_rows] == [30]
    # Neither profile sees legacy chat_id=0 rows!
    assert all(r["outlook_email"] != "legacy_web@outlook.com" for r in p1_rows)
    assert all(r["outlook_email"] != "legacy_web@outlook.com" for r in p2_rows)

    # 3. Both profiles can add the EXACT SAME email independently without overwriting each other
    shared_p1 = db.upsert_vault_entry(
        outlook_email="shared@outlook.com",
        outlook_password="pass_p1",
        tt_login="tt_p1",
        chat_id=111111,
    )
    shared_p2 = db.upsert_vault_entry(
        outlook_email="SHARED@outlook.com",
        outlook_password="pass_p2",
        tt_login="tt_p2",
        chat_id=222222,
    )
    assert shared_p1["id"] != shared_p2["id"]
    assert shared_p1["tt_login"] == "tt_p1"
    assert shared_p2["tt_login"] == "tt_p2"
    assert len(db.get_active_vault_accounts(111111)) == 2
    assert len(db.get_active_vault_accounts(222222)) == 2

    # 4. Cross-profile get / mark_used / delete are blocked by chat_id ownership check
    assert db.get_vault_entry(shared_p1["id"], chat_id=222222) is None
    assert db.get_vault_entry(shared_p1["id"], chat_id=111111) is not None
    assert db.mark_vault_account_used(shared_p1["id"], chat_id=222222) is False
    assert db.delete_vault_entry(shared_p1["id"], chat_id=222222) is False

    # 5. Clearing Profile 2 does NOT affect Profile 1 or legacy chat_id=0
    cleared = db.clear_active_vault_accounts(222222)
    assert cleared == 2
    assert len(db.get_active_vault_accounts(222222)) == 0
    assert len(db.get_active_vault_accounts(111111)) == 2
    assert len(db.get_active_vault_accounts(0)) == 1

    # 6. Profile stats are strictly scoped per profile
    db.upsert_account(
        email="shared@outlook.com",
        display_name="Shared",
        refresh_token="rt_shared",
        access_token="at_shared",
        token_expires_at=9999999999,
    )
    s1 = db.get_profile_stats(111111)
    s2 = db.get_profile_stats(222222)
    assert s1["active"] == 2 and s1["ready_active"] == 1 and s1["deleted"] == 0
    assert s2["active"] == 0 and s2["deleted"] == 2

    # 7. TelegramBot scoped search cancellation & scoped export (no cross-profile leak)
    bot = TelegramBot("dummy-token", db_path=db_path)
    bot.active_searches.add((111111, shared_p1["id"]))
    bot.active_searches.add((222222, shared_p2["id"]))
    bot._cancel_profile_searches(222222)
    assert (222222, shared_p2["id"]) in bot.cancelled_searches
    assert (111111, shared_p1["id"]) not in bot.cancelled_searches

    sent_messages: list[tuple[int, str]] = []
    sent_docs: list[tuple[int, str, str]] = []

    async def fake_send_message(chat_id, text, **kwargs):
        sent_messages.append((int(chat_id), text))
        return {"ok": True}

    async def fake_send_document(chat_id, filename, content, caption=""):
        sent_docs.append((int(chat_id), filename, content))
        return {"ok": True}

    bot.send_message = fake_send_message  # type: ignore[method-assign]
    bot.send_document = fake_send_document  # type: ignore[method-assign]

    # Profile 2 has only deleted entries -> export must NOT leak Profile 1's tokens!
    asyncio.run(bot.export_accounts_txt(222222, profile_id=222222))
    assert len(sent_docs) == 0
    assert len(sent_messages) == 1 and "222222" in sent_messages[0][1]

    # Profile 1 exports ONLY its own tokens
    asyncio.run(bot.export_accounts_txt(111111, profile_id=111111))
    assert len(sent_docs) == 1
    assert sent_docs[0][0] == 111111
    assert sent_docs[0][1] == "outlook_tokens_111111.txt"
    assert "shared@outlook.com" in sent_docs[0][2]
    assert "p2_only@outlook.com" not in sent_docs[0][2]


def test_separate_login_pass_copy_and_retry_button(tmp_path: Path):
    """Verify separate Login/Password copying and 'Повторить' skipping already-delivered codes."""
    import asyncio
    import time

    from db import Database
    from telegram_bot import TelegramBot, _build_copy_rows, _format_cred_lines

    # 1. Credential formatting & copy buttons keep login and password strictly separate
    cred_html = _format_cred_lines(tt_u="tt_user1", tt_p="Secr3t!Pass", email="box@outlook.com")
    assert "<code>tt_user1</code>" in cred_html
    assert "<code>Secr3t!Pass</code>" in cred_html
    assert "tt_user1:Secr3t!Pass" not in cred_html

    rows = _build_copy_rows(
        code="654321", tt_u="tt_user1", tt_p="Secr3t!Pass", email="box@outlook.com"
    )
    assert len(rows) == 2
    assert rows[0][0]["copy_text"]["text"] == "654321"
    assert rows[1][0]["text"] == "📋 Логин" and rows[1][0]["copy_text"]["text"] == "tt_user1"
    assert rows[1][1]["text"] == "📋 Пароль" and rows[1][1]["copy_text"]["text"] == "Secr3t!Pass"

    # Fallback to email + outlook_pass when TikTok creds are empty
    rows_fb = _build_copy_rows(email="box@outlook.com", outlook_pass="MailPass1")
    assert rows_fb[0][0]["copy_text"]["text"] == "box@outlook.com"
    assert rows_fb[0][1]["copy_text"]["text"] == "MailPass1"

    # 2. End-to-end code delivery -> retry skips code #1 and catches code #2
    db_path = tmp_path / "retry_test.db"
    db = Database(db_path)
    v = db.upsert_vault_entry(
        outlook_email="retrybox@outlook.com",
        outlook_password="mp1",
        tt_login="tt_user1",
        tt_password="Secr3t!Pass",
        chat_id=555555,
    )
    vid = v["id"]

    bot = TelegramBot("dummy-token", db_path=db_path)
    sent: list[dict] = []
    now = time.time()

    msg1 = {
        "id": "m1",
        "subject": "111111 is your 6-digit code",
        "sender": "noreply@tiktok.com",
        "received_ts": now - 10,
        "preview": "TikTok verification code 111111",
        "body_text": "TikTok verification code 111111",
    }
    msg2 = {
        "id": "m2",
        "subject": "222222 is your 6-digit code",
        "sender": "noreply@tiktok.com",
        "received_ts": now + 2,
        "preview": "TikTok verification code 222222",
        "body_text": "TikTok verification code 222222",
    }
    mailbox_state = {"msgs": [msg1], "polls": 0}

    async def fake_ensure_access(chat_id, entry, top=6):
        return "tok", "graph", list(mailbox_state["msgs"]), ""

    async def fake_fetch_once(access_token, protocol="graph", top=6):
        mailbox_state["polls"] += 1
        # Simulate new code #2 arriving on the very first poll tick!
        if mailbox_state["polls"] == 1:
            mailbox_state["msgs"] = [msg2, msg1]
        return "ok", list(mailbox_state["msgs"])

    async def fake_send_message(chat_id, text, parse_mode=None, reply_to_message_id=None, reply_markup=None):
        sent.append({"chat_id": chat_id, "text": text, "reply_markup": reply_markup})
        return {"ok": True, "result": {"message_id": len(sent)}}

    async def fake_edit_message(chat_id, message_id, text, parse_mode=None, reply_markup=None):
        return {"ok": True}

    bot._ensure_working_access = fake_ensure_access  # type: ignore[method-assign]
    bot._fetch_messages_once = fake_fetch_once  # type: ignore[method-assign]
    bot.send_message = fake_send_message  # type: ignore[method-assign]
    bot.edit_message = fake_edit_message  # type: ignore[method-assign]

    # First click -> receives 111111 with "🔁 Повторить" button pointing to retry:{vid}:111111
    asyncio.run(bot.request_code_for_vault(555555, vid, profile_id=555555))
    code1_msg = next(m for m in sent if "111111" in m["text"])
    assert "<code>tt_user1</code>" in code1_msg["text"]
    assert "<code>Secr3t!Pass</code>" in code1_msg["text"]
    assert "tt_user1:Secr3t!Pass" not in code1_msg["text"]
    flat_btns = [b for row in code1_msg["reply_markup"]["inline_keyboard"] for b in row]
    retry_btn = next(b for b in flat_btns if b.get("text") == "🔁 Повторить")
    assert retry_btn["callback_data"] == f"retry:{vid}:111111"

    # Clicking "🔁 Повторить" skips already-delivered 111111 and waits for brand-new 222222!
    sent.clear()
    asyncio.run(
        bot.request_code_for_vault(555555, vid, profile_id=555555, skip_prev_code="111111")
    )
    code2_msg = next(m for m in sent if "222222" in m["text"])
    assert "111111" not in code2_msg["text"]
    flat_btns2 = [b for row in code2_msg["reply_markup"]["inline_keyboard"] for b in row]
    retry_btn2 = next(b for b in flat_btns2 if b.get("text") == "🔁 Повторить")
    assert retry_btn2["callback_data"] == f"retry:{vid}:222222"


