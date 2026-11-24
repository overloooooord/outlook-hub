"""Headless browser automation for Microsoft OAuth2 login flow.

Automates 'Step 1': taking Outlook credentials, 2FA TOTP secret, and/or
recovery email credentials (Firstmail IMAP) to complete Microsoft login,
consent approval, and redirect back to Outlook Hub callback.
"""

from __future__ import annotations

import logging
import os
import shutil
import time
from typing import Any, Callable

import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

try:
    from playwright_stealth.stealth import Stealth
    _HAS_STEALTH = True
except ImportError:
    _HAS_STEALTH = False

from vault_utils import generate_totp
from imap_client import fetch_recovery_messages

logger = logging.getLogger("outlook_hub.auto_login")


def _find_chrome_executable() -> str | None:
    for candidate in ("/usr/bin/google-chrome", "/usr/bin/google-chrome-stable", "/usr/bin/chromium", "/usr/bin/chromium-browser"):
        if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
            return candidate
    return shutil.which("google-chrome") or shutil.which("chromium")


def auto_login_account(
    vault_entry: dict[str, Any],
    *,
    base_url: str = "http://127.0.0.1:8765",
    headless: bool = True,
    timeout_sec: int = 70,
    on_progress: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    """Execute automated login for a single vault account."""
    def log(msg: str) -> None:
        logger.info(f"[{vault_entry.get('outlook_email')}] {msg}")
        if on_progress:
            try:
                on_progress(msg)
            except Exception:
                pass

    outlook_email = str(vault_entry.get("outlook_email") or "").strip()
    outlook_password = str(vault_entry.get("outlook_password") or "").strip()
    recovery_email = str(vault_entry.get("recovery_email") or "").strip()
    recovery_password = str(vault_entry.get("recovery_password") or "").strip()
    totp_secret = str(vault_entry.get("totp_secret") or "").strip()

    if not outlook_email:
        raise ValueError("outlook_email отсутствует в записи хранилища")
    if not outlook_password:
        raise ValueError("outlook_password отсутствует в записи хранилища")

    log("Получение ссылки авторизации OAuth...")
    resp = requests.post(
        f"{base_url.rstrip('/')}/api/oauth/start",
        json={"login_hint": outlook_email, "prompt": "login"},
        timeout=15,
    )
    if resp.status_code >= 400:
        raise RuntimeError(f"Ошибка вызова /api/oauth/start: {resp.text}")

    data = resp.json()
    auth_url = data["authorize_url"]
    state = data["state"]

    chrome_exec = _find_chrome_executable()
    log(f"Запуск браузера (headless={headless})...")

    with sync_playwright() as p:
        launch_kwargs: dict[str, Any] = {
            "headless": headless,
            "args": [
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ],
        }
        if chrome_exec:
            launch_kwargs["executable_path"] = chrome_exec

        browser = p.chromium.launch(**launch_kwargs)
        try:
            context = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1280, "height": 800},
            )
            page = context.new_page()
            if _HAS_STEALTH:
                Stealth().apply_stealth_sync(page)

            intercepted_codes: list[str] = []

            def _on_request(req) -> None:
                u = req.url
                # Only intercept nativeclient redirects (which don't hit our local /callback directly).
                # Local /callback requests are already handled directly by the browser navigating to them!
                if "code=" in u and "nativeclient" in u:
                    from urllib.parse import parse_qs, urlparse
                    qs = parse_qs(urlparse(u).query or urlparse(u).fragment)
                    code_val = (qs.get("code") or [""])[0]
                    if code_val and code_val not in intercepted_codes:
                        intercepted_codes.append(code_val)

            page.on("request", _on_request)

            log("Переход на страницу входа Microsoft...")
            page.goto(auth_url, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(2000)

            # 1. Email input (if present)
            email_inp = page.query_selector('input[type="email"], input[name="loginfmt"]')
            if email_inp and email_inp.is_visible():
                log("Ввод логина...")
                email_inp.fill(outlook_email)
                page.wait_for_timeout(300)
                btn = page.query_selector('button[type="submit"], input[type="submit"], #idSIButton9')
                if btn:
                    btn.click()
                page.wait_for_timeout(2500)
            else:
                tile = page.query_selector(f'div[data-test-id="{outlook_email}"]')
                if tile and tile.is_visible():
                    tile.click()
                    page.wait_for_timeout(2500)

            # 2. Password input (Microsoft sometimes pushes "Verify your email" / passwordless first)
            log("Ввод пароля...")
            pwd_inp = None
            clicked_use_password = False
            clicked_other_ways = False

            for _ in range(20):
                pwd_inp = page.query_selector('input[type="password"], input[name="passwd"]')
                if pwd_inp and pwd_inp.is_visible():
                    break

                # If Microsoft offered email-code / passwordless login first, switch to password!
                if not clicked_use_password:
                    for txt in (
                        "Use your password",
                        "Sign in with your password",
                        "Use my password",
                        "Использовать пароль",
                        "Войти с помощью пароля",
                    ):
                        loc = page.get_by_text(txt)
                        if loc.count() and loc.first.is_visible():
                            log(f"Переключение на вход по паролю ({txt})...")
                            loc.first.click()
                            clicked_use_password = True
                            page.wait_for_timeout(1200)
                            break

                # If Microsoft hid "Use your password" behind "Other ways to sign in"
                if not clicked_use_password and not clicked_other_ways:
                    for txt in ("Other ways to sign in", "Другие способы входа"):
                        loc = page.get_by_text(txt)
                        if loc.count() and loc.first.is_visible():
                            log(f"Открытие других способов входа ({txt})...")
                            loc.first.click()
                            clicked_other_ways = True
                            page.wait_for_timeout(1200)
                            break

                page.wait_for_timeout(500)

            if pwd_inp and pwd_inp.is_visible():
                pwd_inp.fill(outlook_password)
                page.wait_for_timeout(300)
                btn = page.query_selector('button[type="submit"], input[type="submit"], #idSIButton9')
                if btn:
                    btn.click()
                page.wait_for_timeout(3000)
            else:
                err_text = page.inner_text("body")
                err_low = err_text.lower()
                if "account doesn't exist" in err_low or "учетная запись не существует" in err_low:
                    raise RuntimeError("Microsoft сообщил: такой аккаунт не существует.")
                # If Microsoft forced recovery email verification without a password option,
                # fall through to Step 3 where Firstmail IMAP verification handles it automatically!
                if not any(k in err_low for k in ("verify your email", "подтвердите адрес", "sent a code", "authenticator")):
                    raise RuntimeError("Поле ввода пароля не появилось.")

            # 3. Handle 2FA / Verification / Consent loop
            start_time = time.time()
            seen_email_verification = False
            last_code_entered = ""

            while time.time() - start_time < timeout_sec:
                if intercepted_codes:
                    code = intercepted_codes[0]
                    log(f"Перехват authorization code ({code[:20]}...) для Thunderbird...")
                    cb_resp = requests.get(
                        f"{base_url.rstrip('/')}/callback",
                        params={"code": code, "state": state},
                        timeout=20,
                    )
                    if cb_resp.status_code >= 400:
                        raise RuntimeError(f"Ошибка обмена кода: {cb_resp.text[:200]}")
                    log("Токены успешно получены и сохранены в базу!")
                    return {"ok": True, "email": outlook_email, "state": state}

                try:
                    cur_url = page.url
                except Exception:
                    page.wait_for_timeout(400)
                    continue

                if "code=" in cur_url and ("nativeclient" in cur_url or "callback" in cur_url or f":{base_url.split(':')[-1]}" in cur_url):
                    if "nativeclient" in cur_url:
                        from urllib.parse import parse_qs, urlparse
                        p_url = urlparse(cur_url)
                        qs = parse_qs(p_url.query or p_url.fragment)
                        code = (qs.get("code") or [""])[0]
                        cb_state = (qs.get("state") or [""])[0] or state
                        if code:
                            log(f"Перехват authorization code ({code[:20]}...) для Thunderbird...")
                            cb_resp = requests.get(
                                f"{base_url.rstrip('/')}/callback",
                                params={"code": code, "state": cb_state},
                                timeout=20,
                            )
                            if cb_resp.status_code >= 400:
                                raise RuntimeError(f"Ошибка обмена кода: {cb_resp.text[:200]}")
                            log("Токены успешно получены и сохранены в базу!")
                            return {"ok": True, "email": outlook_email, "state": state}

                    log("Успешное перенаправление на callback!")
                    # Wait for FastAPI /callback to finish token exchange instead of closing browser early
                    cb_deadline = time.time() + 18
                    while time.time() < cb_deadline:
                        try:
                            st_resp = requests.get(
                                f"{base_url.rstrip('/')}/api/oauth/status/{state}",
                                timeout=5,
                            )
                            if st_resp.status_code == 200:
                                st_data = st_resp.json()
                                if st_data.get("status") == "completed":
                                    log("Аккаунт успешно авторизован и добавлен в систему!")
                                    return {"ok": True, "email": outlook_email, "state": state}
                                if st_data.get("status") == "error":
                                    raise RuntimeError(f"Ошибка в callback: {st_data.get('error')}")
                        except RuntimeError:
                            raise
                        except Exception:
                            pass
                        try:
                            body_text = page.inner_text("body")
                            if "Аккаунт подключён" in body_text or "✔" in body_text:
                                log("Аккаунт успешно авторизован и добавлен в систему!")
                                return {"ok": True, "email": outlook_email, "state": state}
                            if "Ошибка" in body_text or "Вход не выполнен" in body_text:
                                raise RuntimeError(f"Ошибка в callback: {body_text[:200]}")
                        except RuntimeError:
                            raise
                        except Exception:
                            pass
                        page.wait_for_timeout(400)
                    raise RuntimeError("Таймаут завершения обмена токена в /callback.")

                try:
                    body_text = page.inner_text("body").lower()
                    title = page.title().lower()
                except Exception:
                    page.wait_for_timeout(400)
                    continue

                # A. Check for incorrect credentials / lock errors
                if "incorrect password" in body_text or "пароль неверен" in body_text:
                    raise RuntimeError("Неверный пароль от Outlook аккаунта.")
                if "account has been locked" in body_text or "аккаунт заблокирован" in body_text:
                    raise RuntimeError("Аккаунт заблокирован Microsoft.")

                # B. Authenticator App / TOTP 2FA
                if "authenticator" in body_text or "authenticator" in title or (
                    "enter code" in title and totp_secret and not seen_email_verification
                ):
                    if not totp_secret:
                        raise RuntimeError("Microsoft требует 2FA код из Authenticator, но TOTP secret отсутствует.")

                    log("Генерация 2FA TOTP кода...")
                    totp_code, rem = generate_totp(totp_secret)
                    if rem < 4:
                        page.wait_for_timeout((rem + 1) * 1000)
                        totp_code, _ = generate_totp(totp_secret)

                    if totp_code != last_code_entered:
                        log(f"Ввод 2FA кода ({totp_code})...")
                        code_inp = None
                        for inp in page.query_selector_all("input"):
                            if (
                                inp.is_visible()
                                and inp.get_attribute("type") in ("text", "tel", "number", None)
                                and inp.get_attribute("id") != "proof-confirmation-email-input"
                            ):
                                code_inp = inp
                                break
                        if code_inp:
                            code_inp.fill(totp_code)
                            last_code_entered = totp_code
                            page.wait_for_timeout(400)
                            btn = (
                                page.query_selector('button[type="submit"]:visible, input[type="submit"]:visible, #idSIButton9')
                                or page.query_selector('button[type="submit"], input[type="submit"]')
                            )
                            if btn:
                                btn.click()
                            page.wait_for_timeout(3500)
                            continue

                # C. Stay signed in?
                if "stay signed in" in body_text or "оставаться в системе" in body_text:
                    log("Подтверждение 'Оставаться в системе'...")
                    btn = page.query_selector('#idSIButton9, button:has-text("Yes"), input[value="Yes"]')
                    if btn:
                        btn.click()
                    page.wait_for_timeout(3000)
                    continue

                # D. Permissions / Consent ("Let this app access your info?")
                if "accept" in body_text or "let this app access your info" in body_text or "разрешения" in body_text or "consent" in cur_url.lower():
                    log("Принятие разрешений приложения...")
                    accept_btn = page.query_selector('button:has-text("Accept"), input[value="Accept"], #idBtn_Accept, button:has-text("Да"), button:has-text("Yes")')
                    if accept_btn:
                        accept_btn.click()
                        # Wait for redirect to callback
                        try:
                            page.wait_for_url(lambda u: "/callback" in u or "code=" in u or "nativeclient" in u, timeout=12000)
                        except Exception:
                            pass
                        page.wait_for_timeout(2000)
                        continue

                # E. Recovery Email verification challenge ("Verify your email" / "Help us protect your account")
                if "verify your email" in body_text or "подтвердите адрес" in body_text or "help us protect" in body_text:
                    seen_email_verification = True
                    log("Обнаружена проверка резервной почты Microsoft...")

                    # Check if there is an input to confirm recovery email address
                    proof_inp = page.query_selector(
                        'input[id="proof-confirmation-email-input"], '
                        'input[name="otcConfirmationText"], input[id="proofConfirmationText"], '
                        'input[name="proofEmail"], input[type="email"]'
                    )
                    if proof_inp and proof_inp.is_visible():
                        if not recovery_email:
                            raise RuntimeError("Microsoft требует подтвердить резервную почту, но recovery_email отсутствует.")
                        log(f"Ввод адреса резервной почты ({recovery_email})...")
                        proof_inp.fill(recovery_email)
                        page.wait_for_timeout(400)
                        btn = page.query_selector('button[type="submit"], input[type="submit"], #idSIButton9')
                        if btn:
                            btn.click()
                        page.wait_for_timeout(3500)
                        continue

                    # Check if recovery option needs to be clicked (e.g. Email radio/button)
                    email_option = page.query_selector('#iProofEmail, [data-value="Email"], div:has-text("Email")')
                    if email_option and email_option.is_visible():
                        email_option.click()
                        page.wait_for_timeout(1500)
                        continue

                # F. Enter verification code sent to recovery email
                if "sent a code" in body_text or "мы отправили код" in body_text or "we've sent a code" in body_text:
                    if not recovery_email or not recovery_password:
                        raise RuntimeError("Microsoft отправил код на резервную почту, но нет данных для Firstmail IMAP.")

                    log(f"Ожидание кода из Firstmail IMAP ({recovery_email})...")
                    code = None
                    for poll_attempt in range(12):
                        page.wait_for_timeout(2500)
                        try:
                            rec_res = fetch_recovery_messages(
                                email_addr=recovery_email,
                                password=recovery_password,
                                limit=4,
                                timeout=10,
                            )
                            for msg in rec_res.get("messages", []):
                                if msg.get("codes"):
                                    code = msg["codes"][0]
                                    break
                            if code and code != last_code_entered:
                                break
                        except Exception as imap_err:
                            logger.warning(f"IMAP poll error: {imap_err}")

                    if not code:
                        raise RuntimeError(f"Код подтверждения не пришёл на {recovery_email} в течение 30 сек.")

                    log(f"Получен код из Firstmail: {code}. Ввод...")
                    code_inp = None
                    for inp in page.query_selector_all("input"):
                        if inp.is_visible() and inp.get_attribute("type") in ("text", "tel", "number", None):
                            code_inp = inp
                            break
                    if code_inp:
                        code_inp.fill(code)
                        last_code_entered = code
                        page.wait_for_timeout(400)
                        btn = page.query_selector('button[type="submit"], input[type="submit"], #idSIButton9')
                        if btn:
                            btn.click()
                        page.wait_for_timeout(3500)
                        continue

                # G. Skip promo screens ("Use Microsoft Authenticator", "Break free from passwords")
                skip_btn = page.query_selector('#iCancel, a:has-text("No thanks"), button:has-text("No thanks"), a:has-text("Skip")')
                if skip_btn and skip_btn.is_visible():
                    log("Пропуск рекламного экрана Microsoft...")
                    skip_btn.click()
                    page.wait_for_timeout(2000)
                    continue

                page.wait_for_timeout(1000)

            raise TimeoutError(f"Превышено время ожидания входа ({timeout_sec}с). Последний URL: {page.url}")

        finally:
            browser.close()


if __name__ == "__main__":
    import argparse
    import sys
    from db import Database
    from vault_utils import parse_account_lines

    parser = argparse.ArgumentParser(description="Авто-вход (Step 1) в аккаунты Microsoft Outlook")
    parser.add_argument("--add", help="Импортировать строку tt:pass:mail:pass:rec:pass:totp и войти")
    parser.add_argument("--email", help="Выполнить вход для конкретного email из базы")
    parser.add_argument("--all", action="store_true", help="Выполнить вход для всех неподключённых аккаунтов")
    parser.add_argument("--headed", action="store_true", help="Запустить видимый браузер (по умолчанию headless)")
    parser.add_argument("--url", default="http://127.0.0.1:8765", help="Базовый URL Outlook Hub")
    parser.add_argument("--db", default=None, help="Путь к outlook_hub.db")
    args = parser.parse_args()

    db_path = args.db or os.path.join(os.path.dirname(__file__), "data", "outlook_hub.db")
    db = Database(db_path)

    targets = []
    if args.add:
        parsed, skipped = parse_account_lines(args.add)
        if not parsed:
            print("Ошибка: строка не распознана.")
            sys.exit(1)
        for p in parsed:
            entry = db.upsert_vault_entry(**p)
            targets.append(entry)
            print(f"Добавлен в базу: {entry['outlook_email']}")
    elif args.email:
        entry = db.get_vault_entry_by_email(args.email) if hasattr(db, "get_vault_entry_by_email") else None
        if not entry:
            all_entries = db.list_vault_entries()
            entry = next((e for e in all_entries if e["outlook_email"].lower() == args.email.lower()), None)
        if not entry:
            print(f"Ошибка: аккаунт {args.email} не найден в базе.")
            sys.exit(1)
        targets.append(entry)
    else:
        all_entries = db.list_vault_entries()
        targets = [e for e in all_entries if not e["oauth_connected"] or args.all]
        if not targets:
            print("Все аккаунты уже подключены к OAuth!")
            sys.exit(0)

    print(f"Запуск авто-входа для {len(targets)} аккаунтов...")
    for idx, t in enumerate(targets, 1):
        print(f"\n[{idx}/{len(targets)}] Вход в {t['outlook_email']}...")
        try:
            res = auto_login_account(
                t,
                base_url=args.url,
                headless=not args.headed,
                on_progress=lambda m: print(f"  -> {m}"),
            )
            print(f"✔ Успешно подключен: {t['outlook_email']}")
        except Exception as exc:
            print(f"✖ Ошибка для {t['outlook_email']}: {exc}")
