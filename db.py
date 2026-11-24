"""SQLite storage layer for Outlook Hub."""

from __future__ import annotations

import os
import sqlite3
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

DEFAULT_SETTINGS: dict[str, str] = {
    "client_id": "",
    "client_secret": "",
    "tenant": "common",
    "poll_interval_sec": "60",
    "auto_sync": "1",
    "messages_per_folder": "30",
    "sync_junk": "1",
}


class Database:
    """Thread-safe lightweight SQLite wrapper for accounts, tokens, and cached messages."""

    def __init__(self, db_path: str | Path) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.init_db()

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(str(self.db_path), timeout=15.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def init_db(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS accounts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    display_name TEXT NOT NULL DEFAULT '',
                    refresh_token TEXT NOT NULL,
                    access_token TEXT NOT NULL DEFAULT '',
                    token_expires_at REAL NOT NULL DEFAULT 0,
                    client_id TEXT NOT NULL DEFAULT '',
                    client_secret TEXT NOT NULL DEFAULT '',
                    tenant TEXT NOT NULL DEFAULT '',
                    protocol TEXT NOT NULL DEFAULT 'graph',
                    last_sync_at REAL NOT NULL DEFAULT 0,
                    last_error TEXT NOT NULL DEFAULT '',
                    created_at REAL NOT NULL
                );

                CREATE TABLE IF NOT EXISTS messages (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    account_id INTEGER NOT NULL,
                    graph_id TEXT NOT NULL,
                    folder TEXT NOT NULL DEFAULT 'inbox',
                    subject TEXT NOT NULL DEFAULT '',
                    sender_name TEXT NOT NULL DEFAULT '',
                    sender_email TEXT NOT NULL DEFAULT '',
                    to_recipients TEXT NOT NULL DEFAULT '',
                    received_at TEXT NOT NULL DEFAULT '',
                    is_read INTEGER NOT NULL DEFAULT 0,
                    has_attachments INTEGER NOT NULL DEFAULT 0,
                    body_preview TEXT NOT NULL DEFAULT '',
                    body_content TEXT,
                    body_content_type TEXT NOT NULL DEFAULT 'text',
                    web_link TEXT NOT NULL DEFAULT '',
                    synced_at REAL NOT NULL,
                    UNIQUE(account_id, graph_id),
                    FOREIGN KEY (account_id) REFERENCES accounts(id) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS idx_messages_received
                    ON messages(received_at DESC);
                CREATE INDEX IF NOT EXISTS idx_messages_account_folder
                    ON messages(account_id, folder, received_at DESC);

                CREATE TABLE IF NOT EXISTS vault_accounts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    outlook_email TEXT NOT NULL COLLATE NOCASE,
                    outlook_password TEXT NOT NULL DEFAULT '',
                    tt_login TEXT NOT NULL DEFAULT '',
                    tt_password TEXT NOT NULL DEFAULT '',
                    recovery_email TEXT NOT NULL DEFAULT '',
                    recovery_password TEXT NOT NULL DEFAULT '',
                    totp_secret TEXT NOT NULL DEFAULT '',
                    notes TEXT NOT NULL DEFAULT '',
                    created_at REAL NOT NULL,
                    status TEXT NOT NULL DEFAULT 'active',
                    chat_id INTEGER NOT NULL DEFAULT 0,
                    UNIQUE(chat_id, outlook_email)
                );
                """
            )
            # Safe schema migration for existing DBs created before IMAP XOAUTH2 support
            cols = {
                row["name"]
                for row in conn.execute("PRAGMA table_info(accounts)").fetchall()
            }
            if "protocol" not in cols:
                conn.execute(
                    "ALTER TABLE accounts ADD COLUMN protocol TEXT NOT NULL DEFAULT 'graph'"
                )

            v_cols = {
                row["name"]
                for row in conn.execute("PRAGMA table_info(vault_accounts)").fetchall()
            }
            if "status" not in v_cols:
                conn.execute(
                    "ALTER TABLE vault_accounts ADD COLUMN status TEXT NOT NULL DEFAULT 'active'"
                )
            if "chat_id" not in v_cols:
                conn.execute(
                    "ALTER TABLE vault_accounts ADD COLUMN chat_id INTEGER NOT NULL DEFAULT 0"
                )

            # Safe migration: move from global UNIQUE(outlook_email) to per-profile UNIQUE(chat_id, outlook_email)
            ddl_row = conn.execute(
                "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'vault_accounts'"
            ).fetchone()
            ddl_norm = " ".join(str(ddl_row["sql"] if ddl_row else "").lower().split())
            if "unique(chat_id, outlook_email)" not in ddl_norm:
                conn.execute("DROP TABLE IF EXISTS vault_accounts_new")
                conn.execute(
                    """
                    CREATE TABLE vault_accounts_new (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        outlook_email TEXT NOT NULL COLLATE NOCASE,
                        outlook_password TEXT NOT NULL DEFAULT '',
                        tt_login TEXT NOT NULL DEFAULT '',
                        tt_password TEXT NOT NULL DEFAULT '',
                        recovery_email TEXT NOT NULL DEFAULT '',
                        recovery_password TEXT NOT NULL DEFAULT '',
                        totp_secret TEXT NOT NULL DEFAULT '',
                        notes TEXT NOT NULL DEFAULT '',
                        created_at REAL NOT NULL,
                        status TEXT NOT NULL DEFAULT 'active',
                        chat_id INTEGER NOT NULL DEFAULT 0,
                        UNIQUE(chat_id, outlook_email)
                    )
                    """
                )
                conn.execute(
                    """
                    INSERT INTO vault_accounts_new (
                        id, outlook_email, outlook_password, tt_login, tt_password,
                        recovery_email, recovery_password, totp_secret, notes, created_at,
                        status, chat_id
                    )
                    SELECT
                        id, outlook_email, outlook_password, tt_login, tt_password,
                        recovery_email, recovery_password, totp_secret, notes, created_at,
                        COALESCE(status, 'active'), COALESCE(chat_id, 0)
                    FROM vault_accounts
                    """
                )
                conn.execute("DROP TABLE vault_accounts")
                conn.execute("ALTER TABLE vault_accounts_new RENAME TO vault_accounts")

            conn.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_vault_chat_status
                    ON vault_accounts(chat_id, status, id DESC)
                """
            )
            for key, default_val in DEFAULT_SETTINGS.items():
                conn.execute(
                    "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
                    (key, default_val),
                )

        # Restrict permissions so only the current OS user can read saved OAuth tokens.
        try:
            os.chmod(self.db_path, 0o600)
        except OSError:
            pass

    # --- Settings ---

    def get_settings(self) -> dict[str, Any]:
        with self._connect() as conn:
            rows = conn.execute("SELECT key, value FROM settings").fetchall()
        raw = dict(DEFAULT_SETTINGS)
        for row in rows:
            raw[row["key"]] = row["value"]
        return {
            "client_id": raw["client_id"].strip(),
            "client_secret": raw["client_secret"].strip(),
            "tenant": (raw["tenant"] or "common").strip(),
            "poll_interval_sec": max(15, int(raw["poll_interval_sec"] or 60)),
            "auto_sync": raw["auto_sync"] in ("1", "true", "True"),
            "messages_per_folder": max(5, min(100, int(raw["messages_per_folder"] or 30))),
            "sync_junk": raw["sync_junk"] in ("1", "true", "True"),
        }

    def update_settings(self, updates: dict[str, Any]) -> dict[str, Any]:
        allowed = set(DEFAULT_SETTINGS.keys())
        with self._connect() as conn:
            for key, value in updates.items():
                if key not in allowed or value is None:
                    continue
                if isinstance(value, bool):
                    str_val = "1" if value else "0"
                else:
                    str_val = str(value).strip()
                conn.execute(
                    """
                    INSERT INTO settings (key, value) VALUES (?, ?)
                    ON CONFLICT(key) DO UPDATE SET value = excluded.value
                    """,
                    (key, str_val),
                )
        return self.get_settings()

    # --- Accounts ---

    @staticmethod
    def _serialize_account(row: sqlite3.Row, include_tokens: bool = False) -> dict[str, Any]:
        data = dict(row)
        data["has_refresh_token"] = bool(data.get("refresh_token"))
        data["has_custom_client_id"] = bool(data.get("client_id"))
        if not include_tokens:
            data.pop("refresh_token", None)
            data.pop("access_token", None)
            data.pop("client_secret", None)
            data.pop("client_id", None)
        return data

    def upsert_account(
        self,
        *,
        email: str,
        display_name: str,
        refresh_token: str,
        access_token: str = "",
        token_expires_at: float = 0.0,
        client_id: str = "",
        client_secret: str = "",
        tenant: str = "",
        protocol: str = "graph",
    ) -> dict[str, Any]:
        norm_email = email.strip().lower()
        norm_proto = protocol if protocol in ("graph", "imap") else "graph"
        now = time.time()
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO accounts (
                    email, display_name, refresh_token, access_token, token_expires_at,
                    client_id, client_secret, tenant, protocol, last_error, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?)
                ON CONFLICT(email) DO UPDATE SET
                    display_name = CASE
                        WHEN excluded.display_name != '' THEN excluded.display_name
                        ELSE accounts.display_name
                    END,
                    refresh_token = excluded.refresh_token,
                    access_token = excluded.access_token,
                    token_expires_at = excluded.token_expires_at,
                    client_id = CASE
                        WHEN excluded.client_id != '' THEN excluded.client_id
                        ELSE accounts.client_id
                    END,
                    client_secret = CASE
                        WHEN excluded.client_secret != '' THEN excluded.client_secret
                        ELSE accounts.client_secret
                    END,
                    tenant = CASE
                        WHEN excluded.tenant != '' THEN excluded.tenant
                        ELSE accounts.tenant
                    END,
                    protocol = excluded.protocol,
                    last_error = ''
                """,
                (
                    norm_email,
                    display_name.strip(),
                    refresh_token.strip(),
                    access_token,
                    token_expires_at,
                    client_id.strip(),
                    client_secret.strip(),
                    tenant.strip(),
                    norm_proto,
                    now,
                ),
            )
            row = conn.execute(
                "SELECT * FROM accounts WHERE email = ?", (norm_email,)
            ).fetchone()
        assert row is not None
        return self._serialize_account(row, include_tokens=False)

    def list_accounts(self, include_tokens: bool = False) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    a.*,
                    COALESCE(SUM(CASE WHEN m.folder = 'inbox' AND m.is_read = 0 THEN 1 ELSE 0 END), 0) AS unread_count,
                    COALESCE(SUM(CASE WHEN m.folder = 'junkemail' AND m.is_read = 0 THEN 1 ELSE 0 END), 0) AS junk_unread_count,
                    COUNT(m.id) AS total_messages
                FROM accounts a
                LEFT JOIN messages m ON m.account_id = a.id
                GROUP BY a.id
                ORDER BY a.id ASC
                """
            ).fetchall()
        return [self._serialize_account(r, include_tokens=include_tokens) for r in rows]

    def get_account(
        self, account_id: int, include_tokens: bool = False
    ) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM accounts WHERE id = ?", (account_id,)
            ).fetchone()
        if not row:
            return None
        return self._serialize_account(row, include_tokens=include_tokens)

    def get_account_by_email(
        self, email: str, include_tokens: bool = False
    ) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM accounts WHERE LOWER(email) = LOWER(?)", (email.strip(),)
            ).fetchone()
        if not row:
            return None
        return self._serialize_account(row, include_tokens=include_tokens)

    def update_account_tokens(
        self,
        account_id: int,
        *,
        access_token: str,
        token_expires_at: float,
        refresh_token: str | None = None,
    ) -> None:
        with self._connect() as conn:
            if refresh_token:
                conn.execute(
                    """
                    UPDATE accounts
                    SET access_token = ?, token_expires_at = ?, refresh_token = ?, last_error = ''
                    WHERE id = ?
                    """,
                    (access_token, token_expires_at, refresh_token, account_id),
                )
            else:
                conn.execute(
                    """
                    UPDATE accounts
                    SET access_token = ?, token_expires_at = ?, last_error = ''
                    WHERE id = ?
                    """,
                    (access_token, token_expires_at, account_id),
                )

    def update_account_sync_status(
        self,
        account_id: int,
        *,
        last_sync_at: float | None = None,
        last_error: str | None = None,
    ) -> None:
        with self._connect() as conn:
            if last_sync_at is not None and last_error is not None:
                conn.execute(
                    "UPDATE accounts SET last_sync_at = ?, last_error = ? WHERE id = ?",
                    (last_sync_at, last_error, account_id),
                )
            elif last_sync_at is not None:
                conn.execute(
                    "UPDATE accounts SET last_sync_at = ?, last_error = '' WHERE id = ?",
                    (last_sync_at, account_id),
                )
            elif last_error is not None:
                conn.execute(
                    "UPDATE accounts SET last_error = ? WHERE id = ?",
                    (last_error, account_id),
                )

    def delete_account(self, account_id: int) -> bool:
        with self._connect() as conn:
            cur = conn.execute("DELETE FROM accounts WHERE id = ?", (account_id,))
        return cur.rowcount > 0

    # --- Messages ---

    def upsert_messages(
        self, account_id: int, folder: str, items: list[dict[str, Any]]
    ) -> int:
        now = time.time()
        with self._connect() as conn:
            for item in items:
                sender = (item.get("from") or {}).get("emailAddress") or {}
                sender_name = str(sender.get("name") or "").strip()
                sender_email = str(sender.get("address") or "").strip()
                to_list = [
                    ((r.get("emailAddress") or {}).get("address") or "")
                    for r in (item.get("toRecipients") or [])
                ]
                to_str = ", ".join(x for x in to_list if x)
                body_content = item.get("bodyContent")
                body_content_type = str(item.get("bodyContentType") or "text").lower()
                conn.execute(
                    """
                    INSERT INTO messages (
                        account_id, graph_id, folder, subject, sender_name, sender_email,
                        to_recipients, received_at, is_read, has_attachments,
                        body_preview, body_content, body_content_type, web_link, synced_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(account_id, graph_id) DO UPDATE SET
                        folder = excluded.folder,
                        subject = excluded.subject,
                        sender_name = excluded.sender_name,
                        sender_email = excluded.sender_email,
                        to_recipients = CASE
                            WHEN excluded.to_recipients != '' THEN excluded.to_recipients
                            ELSE messages.to_recipients
                        END,
                        received_at = excluded.received_at,
                        is_read = excluded.is_read,
                        has_attachments = excluded.has_attachments,
                        body_preview = excluded.body_preview,
                        body_content = COALESCE(excluded.body_content, messages.body_content),
                        body_content_type = CASE
                            WHEN excluded.body_content IS NOT NULL THEN excluded.body_content_type
                            ELSE messages.body_content_type
                        END,
                        web_link = excluded.web_link,
                        synced_at = excluded.synced_at
                    """,
                    (
                        account_id,
                        item["id"],
                        folder,
                        str(item.get("subject") or "(Без темы)"),
                        sender_name,
                        sender_email,
                        to_str,
                        str(item.get("receivedDateTime") or ""),
                        1 if item.get("isRead") else 0,
                        1 if item.get("hasAttachments") else 0,
                        str(item.get("bodyPreview") or ""),
                        body_content,
                        body_content_type,
                        str(item.get("webLink") or ""),
                        now,
                    ),
                )
        return len(items)

    def update_message_body(
        self,
        message_id: int,
        *,
        body_content: str,
        body_content_type: str,
        to_recipients: str | None = None,
        is_read: bool | None = None,
    ) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                UPDATE messages
                SET body_content = ?,
                    body_content_type = ?,
                    to_recipients = COALESCE(?, to_recipients),
                    is_read = COALESCE(?, is_read)
                WHERE id = ?
                """,
                (
                    body_content,
                    body_content_type,
                    to_recipients,
                    None if is_read is None else (1 if is_read else 0),
                    message_id,
                ),
            )

    def list_messages(
        self,
        *,
        account_id: int | None = None,
        folder: str | None = None,
        unread_only: bool = False,
        search: str | None = None,
        limit: int = 120,
        offset: int = 0,
    ) -> list[dict[str, Any]]:
        clauses: list[str] = []
        params: list[Any] = []

        if account_id is not None:
            clauses.append("m.account_id = ?")
            params.append(account_id)
        if folder and folder != "all":
            clauses.append("m.folder = ?")
            params.append(folder)
        if unread_only:
            clauses.append("m.is_read = 0")
        if search and search.strip():
            like = f"%{search.strip()}%"
            clauses.append(
                "(m.subject LIKE ? OR m.sender_name LIKE ? OR m.sender_email LIKE ? "
                "OR m.body_preview LIKE ? OR a.email LIKE ?)"
            )
            params.extend([like, like, like, like, like])

        where_sql = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        sql = f"""
            SELECT
                m.id, m.account_id, m.graph_id, m.folder, m.subject,
                m.sender_name, m.sender_email, m.to_recipients, m.received_at,
                m.is_read, m.has_attachments, m.body_preview, m.web_link,
                a.email AS account_email, a.display_name AS account_display_name
            FROM messages m
            JOIN accounts a ON a.id = m.account_id
            {where_sql}
            ORDER BY m.received_at DESC, m.id DESC
            LIMIT ? OFFSET ?
        """
        params.extend([limit, offset])

        with self._connect() as conn:
            rows = conn.execute(sql, params).fetchall()
        result = []
        for row in rows:
            item = dict(row)
            item["is_read"] = bool(item["is_read"])
            item["has_attachments"] = bool(item["has_attachments"])
            result.append(item)
        return result

    def get_message(self, message_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT
                    m.*,
                    a.email AS account_email,
                    a.display_name AS account_display_name
                FROM messages m
                JOIN accounts a ON a.id = m.account_id
                WHERE m.id = ?
                """,
                (message_id,),
            ).fetchone()
        if not row:
            return None
        item = dict(row)
        item["is_read"] = bool(item["is_read"])
        item["has_attachments"] = bool(item["has_attachments"])
        return item

    # --- Vault (account credentials + TOTP cards for manual login & OAuth onboarding) ---

    def upsert_vault_entry(
        self,
        *,
        outlook_email: str,
        outlook_password: str = "",
        tt_login: str = "",
        tt_password: str = "",
        recovery_email: str = "",
        recovery_password: str = "",
        totp_secret: str = "",
        notes: str = "",
        chat_id: int = 0,
    ) -> dict[str, Any]:
        norm_email = outlook_email.strip().lower()
        owner_id = int(chat_id or 0)
        now = time.time()
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO vault_accounts (
                    outlook_email, outlook_password, tt_login, tt_password,
                    recovery_email, recovery_password, totp_secret, notes, created_at,
                    status, chat_id
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', ?)
                ON CONFLICT(chat_id, outlook_email) DO UPDATE SET
                    status = 'active',
                    outlook_password = CASE
                        WHEN excluded.outlook_password != '' THEN excluded.outlook_password
                        ELSE vault_accounts.outlook_password
                    END,
                    tt_login = CASE
                        WHEN excluded.tt_login != '' THEN excluded.tt_login
                        ELSE vault_accounts.tt_login
                    END,
                    tt_password = CASE
                        WHEN excluded.tt_password != '' THEN excluded.tt_password
                        ELSE vault_accounts.tt_password
                    END,
                    recovery_email = CASE
                        WHEN excluded.recovery_email != '' THEN excluded.recovery_email
                        ELSE vault_accounts.recovery_email
                    END,
                    recovery_password = CASE
                        WHEN excluded.recovery_password != '' THEN excluded.recovery_password
                        ELSE vault_accounts.recovery_password
                    END,
                    totp_secret = CASE
                        WHEN excluded.totp_secret != '' THEN excluded.totp_secret
                        ELSE vault_accounts.totp_secret
                    END,
                    notes = CASE
                        WHEN excluded.notes != '' THEN excluded.notes
                        ELSE vault_accounts.notes
                    END
                """,
                (
                    norm_email,
                    outlook_password.strip(),
                    tt_login.strip(),
                    tt_password.strip(),
                    recovery_email.strip(),
                    recovery_password.strip(),
                    totp_secret.strip(),
                    notes.strip(),
                    now,
                    owner_id,
                ),
            )
            row = conn.execute(
                """
                SELECT
                    v.*,
                    a.id AS connected_account_id,
                    a.last_error AS connected_error
                FROM vault_accounts v
                LEFT JOIN accounts a ON LOWER(a.email) = LOWER(v.outlook_email)
                WHERE v.chat_id = ? AND v.outlook_email = ?
                """,
                (owner_id, norm_email),
            ).fetchone()
        assert row is not None
        item = dict(row)
        item["oauth_connected"] = item["connected_account_id"] is not None
        return item

    def list_vault_entries(self, chat_id: int | None = None) -> list[dict[str, Any]]:
        with self._connect() as conn:
            if chat_id is not None:
                rows = conn.execute(
                    """
                    SELECT
                        v.*,
                        a.id AS connected_account_id,
                        a.last_error AS connected_error
                    FROM vault_accounts v
                    LEFT JOIN accounts a ON LOWER(a.email) = LOWER(v.outlook_email)
                    WHERE v.chat_id = ?
                    ORDER BY v.id DESC
                    """,
                    (int(chat_id),),
                ).fetchall()
            else:
                rows = conn.execute(
                    """
                    SELECT
                        v.*,
                        a.id AS connected_account_id,
                        a.last_error AS connected_error
                    FROM vault_accounts v
                    LEFT JOIN accounts a ON LOWER(a.email) = LOWER(v.outlook_email)
                    ORDER BY v.id ASC
                    """
                ).fetchall()
        result = []
        for row in rows:
            item = dict(row)
            item["oauth_connected"] = item["connected_account_id"] is not None
            result.append(item)
        return result

    def get_vault_entry(
        self, entry_id: int, chat_id: int | None = None
    ) -> dict[str, Any] | None:
        with self._connect() as conn:
            if chat_id is not None:
                row = conn.execute(
                    "SELECT * FROM vault_accounts WHERE id = ? AND chat_id = ?",
                    (entry_id, int(chat_id)),
                ).fetchone()
            else:
                row = conn.execute(
                    "SELECT * FROM vault_accounts WHERE id = ?", (entry_id,)
                ).fetchone()
        return dict(row) if row else None

    def delete_vault_entry(self, entry_id: int, chat_id: int | None = None) -> bool:
        with self._connect() as conn:
            if chat_id is not None:
                cur = conn.execute(
                    "DELETE FROM vault_accounts WHERE id = ? AND chat_id = ?",
                    (entry_id, int(chat_id)),
                )
            else:
                cur = conn.execute(
                    "DELETE FROM vault_accounts WHERE id = ?", (entry_id,)
                )
        return cur.rowcount > 0

    def get_active_vault_accounts(self, chat_id: int | None = None) -> list[dict[str, Any]]:
        with self._connect() as conn:
            if chat_id is not None:
                cur = conn.execute(
                    "SELECT * FROM vault_accounts WHERE chat_id = ? AND status = 'active' ORDER BY id DESC",
                    (int(chat_id),),
                )
            else:
                cur = conn.execute(
                    "SELECT * FROM vault_accounts WHERE status = 'active' ORDER BY id DESC"
                )
            return [dict(r) for r in cur.fetchall()]

    def mark_vault_account_used(self, entry_id: int, chat_id: int | None = None) -> bool:
        with self._connect() as conn:
            if chat_id is not None:
                cur = conn.execute(
                    "UPDATE vault_accounts SET status = 'used' WHERE id = ? AND chat_id = ?",
                    (entry_id, int(chat_id)),
                )
            else:
                cur = conn.execute(
                    "UPDATE vault_accounts SET status = 'used' WHERE id = ?", (entry_id,)
                )
        return cur.rowcount > 0

    def clear_active_vault_accounts(self, chat_id: int | None = None) -> int:
        with self._connect() as conn:
            if chat_id is not None:
                cur = conn.execute(
                    "UPDATE vault_accounts SET status = 'deleted' WHERE chat_id = ? AND status = 'active'",
                    (int(chat_id),),
                )
            else:
                cur = conn.execute(
                    "UPDATE vault_accounts SET status = 'deleted' WHERE status = 'active'"
                )
        return cur.rowcount

    def get_profile_stats(self, chat_id: int) -> dict[str, Any]:
        owner_id = int(chat_id)
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT
                    COUNT(*) AS total,
                    SUM(CASE WHEN v.status = 'active' THEN 1 ELSE 0 END) AS active,
                    SUM(CASE WHEN v.status = 'used' THEN 1 ELSE 0 END) AS used,
                    SUM(CASE WHEN v.status = 'deleted' THEN 1 ELSE 0 END) AS deleted,
                    SUM(
                        CASE WHEN v.status = 'active' AND a.refresh_token IS NOT NULL AND a.refresh_token != ''
                        THEN 1 ELSE 0 END
                    ) AS ready_active,
                    SUM(
                        CASE WHEN a.refresh_token IS NOT NULL AND a.refresh_token != ''
                        THEN 1 ELSE 0 END
                    ) AS ready_total
                FROM vault_accounts v
                LEFT JOIN accounts a ON LOWER(a.email) = LOWER(v.outlook_email)
                WHERE v.chat_id = ?
                """,
                (owner_id,),
            ).fetchone()
        return {
            "profile_id": owner_id,
            "total": int((row["total"] if row else 0) or 0),
            "active": int((row["active"] if row else 0) or 0),
            "used": int((row["used"] if row else 0) or 0),
            "deleted": int((row["deleted"] if row else 0) or 0),
            "ready_active": int((row["ready_active"] if row else 0) or 0),
            "ready_total": int((row["ready_total"] if row else 0) or 0),
        }

    def get_stats(self) -> dict[str, Any]:
        with self._connect() as conn:
            acc = conn.execute(
                """
                SELECT
                    COUNT(*) AS total_accounts,
                    SUM(CASE WHEN last_error != '' THEN 1 ELSE 0 END) AS error_accounts
                FROM accounts
                """
            ).fetchone()
            msg = conn.execute(
                """
                SELECT
                    COUNT(*) AS total_messages,
                    SUM(CASE WHEN folder = 'inbox' AND is_read = 0 THEN 1 ELSE 0 END) AS unread_inbox,
                    SUM(CASE WHEN folder = 'junkemail' THEN 1 ELSE 0 END) AS total_junk,
                    SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) AS unread_all
                FROM messages
                """
            ).fetchone()
            vault = conn.execute("SELECT COUNT(*) AS total_vault FROM vault_accounts").fetchone()
        return {
            "total_accounts": int(acc["total_accounts"] or 0),
            "error_accounts": int(acc["error_accounts"] or 0),
            "total_messages": int(msg["total_messages"] or 0),
            "unread_inbox": int(msg["unread_inbox"] or 0),
            "total_junk": int(msg["total_junk"] or 0),
            "unread_all": int(msg["unread_all"] or 0),
            "total_vault": int(vault["total_vault"] or 0),
        }
