# Outlook Hub

Единый веб-клиент для одновременной работы с 20+ аккаунтами Outlook через Microsoft OAuth2 (`refresh_token`): **Graph API**, **IMAP XOAUTH2**, **Device Code Flow**, **база аккаунтов с живыми 2FA (TOTP)** и **встроенным чтением резервных почт Firstmail по IMAP**.

## Возможности
- **Microsoft Graph API + IMAP XOAUTH2 (`outlook.office365.com:993`)** — автоматически определяет протокол при импорте токенов, сам обновляет `access_token` и сохраняет ротируемые `refresh_token`.
- **⚡ 100% Автоматизация 1-го шага (Step 1 Auto-Login)**:
  - Автоматический вход через headless Chrome со stealth-маскировкой (`auto_login.py`).
  - Сам вводит логин и пароль Outlook.
  - При запросе 2FA автоматически генерирует и подставляет код из TOTP-ключа.
  - При запросе подтверждения резервной почты — вводит адрес Firstmail, подключается по IMAP к `imap.firstmail.ltd:993`, забирает 6-значный код безопасности и вводит его.
  - Автоматически принимает разрешения (Consent) и сохраняет полученный `refresh_token`.
  - Запускается в 1 клик прямо из веба («⚡ Авто-вход для всех») или из консоли (`python3 auto_login.py`).
- **Четыре способа подключения ящиков:**
  1. **⚡ Авто-вход (Step 1)** — полностью автоматически без единого клика руками в браузере.
  2. **Вход Microsoft (OAuth2 PKCE)** — с кнопкой «Скопировать ссылку (для Инкогнито)».
  3. **Код устройства (Device Code Flow)** — работает вообще без Redirect URI (удобно на VPS по IP, с телефона или в антидетекте).
  4. **Импорт `refresh_token`** — одиночный токен или сразу пачка строк `email:pass:refresh_token:client_id`.
- **База аккаунтов + живые 2FA-коды (RFC 6238 TOTP)** + кнопка **«Забрать код (резерв)»** (читает `imap.firstmail.ltd:993` / Rambler / Mail.ru / Yandex и сразу вытаскивает коды подтверждения Microsoft).
- **Готов к запуску локально и на VPS** (встроенная защита паролем через `OUTLOOK_HUB_PASSWORD`).

## Локальный запуск

```bash
pip install -r requirements.txt
playwright install chromium
cp .env.example .env   # пропиши TELEGRAM_BOT_TOKEN и OUTLOOK_HUB_CLIENT_ID
```

1. **Запуск локального OAuth2 Hub (`http://127.0.0.1:8765`):**
```bash
python3 app.py
```

2. **Запуск Telegram-бота выдачи кодов (с изолированными профилями по Telegram ID):**
```bash
python3 telegram_bot.py
```

### Возможности Telegram-бота (`telegram_bot.py`)
- **Изолированные профили по Telegram ID** (`/profile`, `/stats`, `/clear`, `/export`) — у каждого Telegram-аккаунта свой независимый список почт, поиск кодов, очистка и экспорт токенов.
- **Мгновенное получение кодов TikTok / OTP через Microsoft Graph API** (~0.4 сек на запрос, работает даже если на почте отключён IMAP).
- **Удобное раздельное копирование в 1 клик** (`copy_text`-кнопки для кода, логина и пароля по отдельности) + кнопка **`🔁 Повторить`** для ожидания следующего свежего кода с того же ящика.
- **Фоновый авто-прогрев OAuth2 токенов** при отправке строк или `.txt` файла с аккаунтами.

## Запуск на VPS (с паролем)

```bash
OUTLOOK_HUB_HOST=0.0.0.0 \
OUTLOOK_HUB_PORT=8765 \
OUTLOOK_HUB_PASSWORD="придумай-сложный-пароль" \
OUTLOOK_HUB_PUBLIC_URL="https://твой-домен.com" \
python3 app.py
```

- Если запускаешь на VPS **без домена (просто по IP)** — `OUTLOOK_HUB_PUBLIC_URL` можно не задавать и добавлять ящики через вкладку **«Код устройства (VPS)»** (`microsoft.com/devicelogin`).
- Для Device Code Flow один раз включи в Entra: **Authentication → Advanced settings → Allow public client flows = Yes → Save**.
