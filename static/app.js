(() => {
  const state = {
    settings: null,
    accounts: [],
    vaultEntries: [],
    vaultSearch: "",
    showVaultSecrets: false,
    recoveryCache: {},
    activeRecoveryEntryId: null,
    stats: null,
    messages: [],
    selectedView: "all-inbox", // all-inbox | all-unread | all-junk | all-everything | account | vault
    selectedAccountId: null,
    selectedMessageId: null,
    activeMessage: null,
    bodyViewMode: "html", // html | text
    searchQuery: "",
    accountFilter: "",
    oauthPollTimer: null,
    uiPollTimer: null,
    totpTickTimer: null,
    lastTotpStep: Math.floor(Date.now() / 30000),
    lastGlobalSync: 0,
  };

  const el = (id) => document.getElementById(id);

  function escapeHtml(str) {
    return String(str ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function accountColor(email) {
    let hash = 0;
    const s = String(email || "").toLowerCase();
    for (let i = 0; i < s.length; i++) {
      hash = (hash * 31 + s.charCodeAt(i)) >>> 0;
    }
    const hue = hash % 360;
    return {
      dot: `hsl(${hue}, 72%, 58%)`,
      bg: `hsla(${hue}, 70%, 50%, 0.14)`,
      fg: `hsl(${hue}, 82%, 76%)`,
      border: `hsla(${hue}, 70%, 55%, 0.35)`,
    };
  }

  function formatShortTime(isoStr) {
    if (!isoStr) return "";
    const d = new Date(isoStr);
    if (Number.isNaN(d.getTime())) return "";
    const now = new Date();
    const sameDay =
      d.getFullYear() === now.getFullYear() &&
      d.getMonth() === now.getMonth() &&
      d.getDate() === now.getDate();
    if (sameDay) {
      return d.toLocaleTimeString("ru-RU", { hour: "2-digit", minute: "2-digit" });
    }
    return d.toLocaleDateString("ru-RU", { day: "numeric", month: "short" });
  }

  function formatFullDate(isoStr) {
    if (!isoStr) return "";
    const d = new Date(isoStr);
    if (Number.isNaN(d.getTime())) return isoStr;
    return d.toLocaleString("ru-RU", {
      day: "2-digit",
      month: "long",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  function htmlToPlainText(htmlStr) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(htmlStr || "", "text/html");
    doc.querySelectorAll("script, style, noscript").forEach((n) => n.remove());
    return (doc.body?.textContent || "").replace(/\n{3,}/g, "\n\n").trim();
  }

  function toast(message, type = "") {
    const box = document.createElement("div");
    box.className = `toast ${type}`;
    box.textContent = message;
    el("toastContainer").appendChild(box);
    setTimeout(() => box.remove(), 3800);
  }

  async function api(url, options = {}) {
    const resp = await fetch(url, {
      headers: { "Content-Type": "application/json", ...(options.headers || {}) },
      ...options,
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || `Ошибка HTTP ${resp.status}`);
    }
    return data;
  }

  async function copyText(text, okMessage = "Скопировано") {
    try {
      await navigator.clipboard.writeText(text);
      toast(okMessage, "ok");
    } catch {
      const ta = document.createElement("textarea");
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      ta.remove();
      toast(okMessage, "ok");
    }
  }

  // --- Data loading ---

  async function loadSettings() {
    const s = await api("/api/settings");
    state.settings = s;
    el("onboardingRedirectUri").textContent = s.redirect_uri;
    el("vaultRedirectUri").textContent = s.redirect_uri;
    el("settingsRedirectUri").textContent = s.redirect_uri;
    el("setClientId").value = s.client_id || "";
    el("setTenant").value = s.tenant || "common";
    el("setClientSecret").value = "";
    el("secretStatusHint").textContent = s.has_client_secret
      ? "Client Secret сейчас сохранён (оставь поле пустым, чтобы не менять)."
      : "Client Secret не задан (для Public client он и не нужен).";
    el("setPollInterval").value = s.poll_interval_sec;
    el("setMessagesPerFolder").value = s.messages_per_folder;
    el("setAutoSync").checked = Boolean(s.auto_sync);
    el("setSyncJunk").checked = Boolean(s.sync_junk);

    el("setupBanner").classList.toggle("hidden", Boolean(s.configured));
    el("vaultSetupBanner").classList.toggle("hidden", Boolean(s.configured));
  }

  async function loadAccounts() {
    const data = await api("/api/accounts");
    state.accounts = data.accounts || [];
    state.stats = data.stats;
    state.lastGlobalSync = data.last_global_sync || 0;
    setSyncSpinner(Boolean(data.syncing));
    renderStats();
    renderAccounts();
  }

  async function loadMessages() {
    const params = new URLSearchParams();
    if (state.selectedView === "account" && state.selectedAccountId) {
      params.set("account_id", String(state.selectedAccountId));
      params.set("folder", "all");
    } else if (state.selectedView === "all-inbox") {
      params.set("folder", "inbox");
    } else if (state.selectedView === "all-unread") {
      params.set("folder", "all");
      params.set("unread_only", "true");
    } else if (state.selectedView === "all-junk") {
      params.set("folder", "junkemail");
    } else {
      params.set("folder", "all");
    }
    if (state.searchQuery.trim()) {
      params.set("search", state.searchQuery.trim());
    }
    params.set("limit", "200");

    const data = await api(`/api/messages?${params.toString()}`);
    state.messages = data.messages || [];
    state.stats = data.stats;
    renderStats();
    renderMessages();
  }

  async function openMessage(messageId, forceRefresh = false) {
    state.selectedMessageId = messageId;
    renderMessages();

    el("readerEmpty").classList.add("hidden");
    el("readerContent").classList.remove("hidden");
    el("readerLoading").classList.remove("hidden");
    el("readerIframe").classList.add("hidden");
    el("readerText").classList.add("hidden");

    const preview = state.messages.find((m) => m.id === messageId);
    if (preview) {
      renderReaderHeader(preview);
    }

    try {
      const qs = forceRefresh ? "?refresh=true" : "";
      const data = await api(`/api/messages/${messageId}${qs}`);
      if (state.selectedMessageId !== messageId) return;
      state.activeMessage = data.message;
      state.bodyViewMode =
        data.message.body_content_type === "html" ? "html" : "text";
      renderReaderHeader(data.message);
      renderReaderBody(data.message);
    } catch (err) {
      el("readerLoading").classList.add("hidden");
      el("readerText").classList.remove("hidden");
      el("readerText").textContent = `Не удалось загрузить тело письма: ${err.message}`;
    }
  }

  // --- Rendering ---

  async function loadVault() {
    const data = await api("/api/vault");
    state.vaultEntries = data.entries || [];
    state.stats = data.stats || state.stats;
    renderStats();
    renderVault();
  }

  function renderStats() {
    const s = state.stats || {
      total_accounts: 0,
      error_accounts: 0,
      total_messages: 0,
      unread_inbox: 0,
      total_junk: 0,
      unread_all: 0,
      total_vault: 0,
    };
    const errText = s.error_accounts ? ` • ${s.error_accounts} с ошибкой` : "";
    const vaultPart = s.total_vault ? ` • база: ${s.total_vault}` : "";
    el("statsSummary").textContent = `${s.total_accounts} почт${vaultPart}${errText}`;
    el("badgeUnreadInbox").textContent = String(s.unread_inbox);
    el("badgeUnreadAll").textContent = String(s.unread_all);
    el("badgeTotalJunk").textContent = String(s.total_junk);
    el("badgeTotalAll").textContent = String(s.total_messages);
    el("badgeTotalVault").textContent = String(s.total_vault || 0);

    if (state.lastGlobalSync > 0) {
      const t = new Date(state.lastGlobalSync * 1000).toLocaleTimeString("ru-RU", {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      });
      el("syncStatusText").textContent = `Обновлено в ${t}`;
    }
  }

  function renderVault() {
    const grid = el("vaultGrid");
    const q = state.vaultSearch.trim().toLowerCase();
    const list = q
      ? state.vaultEntries.filter((e) =>
          [e.outlook_email, e.tt_login, e.recovery_email]
            .join(" ")
            .toLowerCase()
            .includes(q)
        )
      : state.vaultEntries;

    if (!list.length) {
      grid.innerHTML = `
        <div class="empty-state">
          <div class="empty-title">${
            state.vaultEntries.length ? "Ничего не найдено" : "База аккаунтов пуста"
          }</div>
          <div class="muted">Нажми «+ Загрузить список», чтобы добавить строки <code>лог:пасс:почта:пасс:резерв:пасс:2фа</code>.</div>
        </div>`;
      return;
    }

    const mask = (v) => (state.showVaultSecrets ? v : "•".repeat(Math.min(12, Math.max(6, v.length))));
    const fmtTotp = (code) =>
      code && code.length === 6 ? `${code.slice(0, 3)} ${code.slice(3)}` : "—";

    grid.innerHTML = "";
    for (const item of list) {
      const color = accountColor(item.outlook_email);
      const card = document.createElement("div");
      card.className = "vault-card";

      const rows = [];
      rows.push({
        label: "Outlook почта",
        display: item.outlook_email,
        raw: item.outlook_email,
      });
      if (item.outlook_password) {
        rows.push({
          label: "Пароль Outlook",
          display: mask(item.outlook_password),
          raw: item.outlook_password,
        });
      }
      if (item.tt_login) {
        rows.push({ label: "TikTok логин", display: item.tt_login, raw: item.tt_login });
      }
      if (item.tt_password) {
        rows.push({
          label: "TikTok пароль",
          display: mask(item.tt_password),
          raw: item.tt_password,
        });
      }
      if (item.recovery_email) {
        rows.push({
          label: "Резервная почта",
          display: item.recovery_email,
          raw: item.recovery_email,
        });
      }
      if (item.recovery_password) {
        rows.push({
          label: "Пароль резервной почты",
          display: mask(item.recovery_password),
          raw: item.recovery_password,
        });
      }

      const rowsHtml = rows
        .map(
          (r, idx) => `
          <div class="vault-row">
            <div class="vault-row-left">
              <div class="vault-label">${escapeHtml(r.label)}</div>
              <div class="vault-val" title="${escapeHtml(r.display)}">${escapeHtml(r.display)}</div>
            </div>
            <button class="btn btn-ghost btn-xs" data-copy-idx="${idx}">Коп</button>
          </div>`
        )
        .join("");

      const totpHtml = item.totp_secret
        ? `
          <div class="totp-box">
            <div>
              <div class="vault-label">2FA код (TOTP)</div>
              <div class="totp-code" data-totp-code="${item.id}">${escapeHtml(fmtTotp(item.totp_code))}</div>
            </div>
            <div style="display:flex;align-items:center;gap:8px;">
              <span class="totp-timer" data-totp-timer="${item.id}">${item.totp_remaining || 30}с</span>
              <button class="btn btn-primary btn-xs" data-act="copy-totp">Коп код</button>
              <button class="btn btn-ghost btn-xs" data-act="copy-secret" title="Копировать Base32 ключ">Ключ</button>
            </div>
          </div>`
        : "";

      const hasRecovery = Boolean(item.recovery_email && item.recovery_password);
      const recState = state.recoveryCache[item.id];
      let recoveryHtml = "";
      if (recState?.loading) {
        recoveryHtml = `<div class="recovery-inline loading">Подключаюсь к ${escapeHtml(item.recovery_email)} по IMAP...</div>`;
      } else if (recState?.error) {
        recoveryHtml = `<div class="recovery-inline error">${escapeHtml(recState.error)}</div>`;
      } else if (recState?.data) {
        const msgs = recState.data.messages || [];
        const latestWithCode = msgs.find((m) => (m.codes || []).length > 0);
        const latestCode = latestWithCode ? latestWithCode.codes[0] : "";
        const latestMsg = latestWithCode || msgs[0];
        if (latestCode) {
          recoveryHtml = `
            <div class="recovery-inline">
              <div class="recovery-code-row">
                <div>
                  <div class="vault-label">Код с резервной почты</div>
                  <div class="recovery-code-val">${escapeHtml(latestCode)}</div>
                </div>
                <div style="display:flex;gap:6px;">
                  <button class="btn btn-primary btn-xs" data-act="copy-rec-code" data-code="${escapeHtml(latestCode)}">Коп код</button>
                  <button class="btn btn-ghost btn-xs" data-act="open-rec-modal">Письма (${msgs.length})</button>
                </div>
              </div>
              <div class="muted small" style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${escapeHtml(latestMsg.subject)}">
                ${escapeHtml(latestMsg.subject)}
              </div>
            </div>`;
        } else if (latestMsg) {
          recoveryHtml = `
            <div class="recovery-inline">
              <div class="recovery-code-row">
                <div class="small" style="min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">
                  ${escapeHtml(latestMsg.subject)}
                </div>
                <button class="btn btn-ghost btn-xs" data-act="open-rec-modal">Письма (${msgs.length})</button>
              </div>
              <div class="muted small">${escapeHtml(latestMsg.preview || "Код не найден — открой письмо")}</div>
            </div>`;
        } else {
          recoveryHtml = `
            <div class="recovery-inline">
              <div class="recovery-code-row">
                <span class="muted small">Писем пока нет (проверено ${escapeHtml(recState.data.imap_host)})</span>
                <button class="btn btn-ghost btn-xs" data-act="open-rec-modal">Открыть</button>
              </div>
            </div>`;
        }
      }

      card.innerHTML = `
        <div class="vault-card-head">
          <span class="account-pill" style="background:${color.bg};color:${color.fg};border-color:${color.border}">
            ${escapeHtml(item.outlook_email)}
          </span>
          <span class="vault-status-pill ${item.oauth_connected ? "connected" : "pending"}">
            ${item.oauth_connected ? "Почта подключена" : "OAuth не подключён"}
          </span>
        </div>
        <div class="vault-rows">${rowsHtml}</div>
        ${totpHtml}
        ${recoveryHtml}
        <div class="vault-card-actions">
          ${
            hasRecovery
              ? `<button class="btn btn-secondary btn-xs" data-act="fetch-recovery">Забрать код (резерв)</button>`
              : ""
          }
          ${
            item.oauth_connected
              ? `<button class="btn btn-primary btn-xs" data-act="open-mail">Письма</button>
                 <button class="btn btn-ghost btn-xs" data-act="auto-login" title="Обновить токен через headless browser">⚡ Авто-вход</button>`
              : `<button class="btn btn-primary btn-xs" data-act="auto-login" style="background:#0284c7;border-color:#0284c7;">⚡ Авто-вход</button>
                 <button class="btn btn-secondary btn-xs" data-act="connect-oauth">Вручную</button>`
          }
          <button class="btn btn-ghost btn-xs" data-act="delete-vault" style="margin-left:auto;color:#f87171;">Удалить</button>
        </div>
      `;

      card.querySelectorAll("button[data-copy-idx]").forEach((btn) => {
        btn.addEventListener("click", () => {
          const r = rows[Number(btn.dataset.copyIdx)];
          copyText(r.raw, `${r.label}: скопировано`);
        });
      });

      card.querySelector("[data-act='auto-login']")?.addEventListener("click", async (e) => {
        const btn = e.currentTarget;
        const origText = btn.textContent;
        btn.disabled = true;
        btn.textContent = "⏳ Вход...";
        toast(`Запущен авто-вход для ${item.outlook_email}... Подождите 15-20 сек.`);
        try {
          await api(`/api/vault/${item.id}/auto-login`, { method: "POST" });
          toast(`✔ ${item.outlook_email} успешно подключен! Загружаем письма...`, "ok");
          await loadAccounts();
          await loadVault();
          await loadStats();
        } catch (err) {
          toast(`Ошибка авто-входа: ${err.message || err}`, "err");
        } finally {
          btn.disabled = false;
          btn.textContent = origText;
        }
      });

      card.querySelector("[data-act='copy-totp']")?.addEventListener("click", () => {
        if (item.totp_code) {
          copyText(item.totp_code, `2FA код ${item.totp_code} скопирован`);
        }
      });

      card.querySelector("[data-act='copy-secret']")?.addEventListener("click", () => {
        copyText(item.totp_secret, "Base32 2FA ключ скопирован");
      });

      card.querySelector("[data-act='fetch-recovery']")?.addEventListener("click", () => {
        fetchRecoveryMail(item);
      });

      card.querySelector("[data-act='copy-rec-code']")?.addEventListener("click", (e) => {
        const code = e.currentTarget.dataset.code;
        if (code) copyText(code, `Код ${code} скопирован`);
      });

      card.querySelector("[data-act='open-rec-modal']")?.addEventListener("click", () => {
        openRecoveryModal(item);
      });

      card.querySelector("[data-act='connect-oauth']")?.addEventListener("click", () => {
        openAddAccountModal(item.outlook_email);
      });

      card.querySelector("[data-act='open-mail']")?.addEventListener("click", () => {
        selectView("account", item.connected_account_id, item.outlook_email);
      });

      card.querySelector("[data-act='delete-vault']")?.addEventListener("click", async () => {
        if (!confirm(`Удалить ${item.outlook_email} из базы аккаунтов?`)) return;
        try {
          await api(`/api/vault/${item.id}`, { method: "DELETE" });
          delete state.recoveryCache[item.id];
          toast(`Удалён ${item.outlook_email}`, "ok");
          await loadVault();
        } catch (err) {
          toast(err.message, "err");
        }
      });

      grid.appendChild(card);
    }
  }

  async function fetchRecoveryMail(item, imapHost = "", refreshModal = false) {
    state.recoveryCache[item.id] = { loading: true, error: "", data: state.recoveryCache[item.id]?.data || null };
    renderVault();
    if (refreshModal && state.activeRecoveryEntryId === item.id) {
      el("recoveryMessagesList").innerHTML = `<div class="reader-loading">Загружаю письма по IMAP...</div>`;
    }
    try {
      const qs = imapHost.trim() ? `?imap_host=${encodeURIComponent(imapHost.trim())}` : "";
      const data = await api(`/api/vault/${item.id}/recovery-mail${qs}`, { method: "POST" });
      state.recoveryCache[item.id] = { loading: false, error: "", data };
      renderVault();

      const withCode = (data.messages || []).find((m) => (m.codes || []).length > 0);
      if (withCode) {
        toast(`Код из ${data.recovery_email}: ${withCode.codes[0]}`, "ok");
      } else {
        toast(`Загружено писем: ${(data.messages || []).length}`, "ok");
      }

      if (refreshModal && state.activeRecoveryEntryId === item.id) {
        renderRecoveryModal(item);
      }
    } catch (err) {
      state.recoveryCache[item.id] = { loading: false, error: err.message, data: null };
      renderVault();
      toast(err.message, "err");
      if (refreshModal && state.activeRecoveryEntryId === item.id) {
        el("recoveryMessagesList").innerHTML = `<div class="oauth-status error">${escapeHtml(err.message)}</div>`;
      }
    }
  }

  function openRecoveryModal(item) {
    state.activeRecoveryEntryId = item.id;
    el("recoveryModalTitle").textContent = `Резервная почта: ${item.recovery_email}`;
    const cached = state.recoveryCache[item.id];
    el("recoveryModalSub").textContent = cached?.data?.imap_host
      ? `IMAP: ${cached.data.imap_host} • для ${item.outlook_email}`
      : `Для ${item.outlook_email}`;
    el("recoveryImapHostInput").value = cached?.data?.imap_host || "";
    openModal("modalRecoveryMail");
    renderRecoveryModal(item);
  }

  function renderRecoveryModal(item) {
    const listEl = el("recoveryMessagesList");
    const cached = state.recoveryCache[item.id];
    if (cached?.loading) {
      listEl.innerHTML = `<div class="reader-loading">Загружаю письма по IMAP...</div>`;
      return;
    }
    if (cached?.error) {
      listEl.innerHTML = `<div class="oauth-status error">${escapeHtml(cached.error)}</div>`;
      return;
    }
    const msgs = cached?.data?.messages || [];
    if (!msgs.length) {
      listEl.innerHTML = `<div class="empty-state"><div class="empty-title">Писем нет</div><div class="muted">Подожди 10–15 секунд после нажатия Send code и нажми «Обновить письма».</div></div>`;
      return;
    }

    listEl.innerHTML = "";
    for (const msg of msgs) {
      const card = document.createElement("div");
      card.className = "recovery-msg-card";
      const codesHtml = (msg.codes || []).length
        ? `<div class="recovery-codes-bar">${msg.codes
            .map((c) => `<button class="code-chip" data-copy-code="${escapeHtml(c)}">Код: ${escapeHtml(c)} (коп)</button>`)
            .join("")}</div>`
        : "";
      card.innerHTML = `
        <div class="recovery-msg-head">
          <span>${escapeHtml(msg.sender || "Неизвестный отправитель")}</span>
          <span>${escapeHtml(msg.date || "")}</span>
        </div>
        <div class="recovery-msg-subj">${escapeHtml(msg.subject)}</div>
        ${codesHtml}
        <pre class="recovery-msg-text">${escapeHtml(msg.text || msg.preview || "")}</pre>
      `;
      card.querySelectorAll("button[data-copy-code]").forEach((btn) => {
        btn.addEventListener("click", () => {
          const code = btn.dataset.copyCode;
          copyText(code, `Код ${code} скопирован`);
        });
      });
      listEl.appendChild(card);
    }
  }

  function renderAccounts() {
    const container = el("accountsList");
    const filter = state.accountFilter.trim().toLowerCase();
    const list = filter
      ? state.accounts.filter((a) => a.email.toLowerCase().includes(filter))
      : state.accounts;

    if (!list.length) {
      container.innerHTML = `<div class="empty-hint">${
        state.accounts.length ? "Ничего не найдено" : "Пока нет подключённых ящиков"
      }</div>`;
      return;
    }

    container.innerHTML = "";
    for (const acc of list) {
      const color = accountColor(acc.email);
      const row = document.createElement("div");
      row.className = `account-row${
        state.selectedView === "account" && state.selectedAccountId === acc.id
          ? " active"
          : ""
      }`;
      row.title = acc.last_error ? `${acc.email}\nОшибка: ${acc.last_error}` : acc.email;

      const dotColor = acc.last_error ? "#ef4444" : color.dot;
      row.innerHTML = `
        <span class="account-dot" style="background:${dotColor}"></span>
        <div class="account-info">
          <div class="account-email">${escapeHtml(acc.email)}</div>
          ${acc.last_error ? `<div class="account-err">${escapeHtml(acc.last_error)}</div>` : ""}
        </div>
        ${
          acc.unread_count > 0
            ? `<span class="badge account-unread">${acc.unread_count}</span>`
            : ""
        }
        <div class="account-actions">
          <button class="tiny-btn" data-act="copy" title="Копировать email">Коп</button>
          <button class="tiny-btn" data-act="sync" title="Синхронизировать">&#8635;</button>
          ${
            acc.last_error
              ? `<button class="tiny-btn" data-act="reauth" title="Войти заново">Вход</button>`
              : ""
          }
          <button class="tiny-btn danger" data-act="delete" title="Удалить">&times;</button>
        </div>
      `;

      row.addEventListener("click", (e) => {
        const btn = e.target.closest("button[data-act]");
        if (btn) {
          e.stopPropagation();
          handleAccountAction(btn.dataset.act, acc);
          return;
        }
        selectView("account", acc.id, acc.email);
      });

      container.appendChild(row);
    }
  }

  async function handleAccountAction(action, acc) {
    if (action === "copy") {
      await copyText(acc.email, `Скопировано: ${acc.email}`);
      return;
    }
    if (action === "reauth") {
      openAddAccountModal(acc.email);
      return;
    }
    if (action === "sync") {
      toast(`Обновляю ${acc.email}...`);
      try {
        const res = await api(`/api/accounts/${acc.id}/sync`, { method: "POST" });
        state.accounts = res.accounts;
        state.stats = res.stats;
        renderStats();
        renderAccounts();
        await loadMessages();
        if (res.result.ok) {
          toast(`${acc.email}: обновлено (${res.result.fetched})`, "ok");
        } else {
          toast(`${acc.email}: ${res.result.error}`, "err");
        }
      } catch (err) {
        toast(err.message, "err");
      }
      return;
    }
    if (action === "delete") {
      if (!confirm(`Удалить аккаунт ${acc.email} и его кэш писем из Outlook Hub?`)) {
        return;
      }
      try {
        await api(`/api/accounts/${acc.id}`, { method: "DELETE" });
        if (state.selectedView === "account" && state.selectedAccountId === acc.id) {
          selectView("all-inbox");
        } else {
          await loadAccounts();
          await loadMessages();
        }
        toast(`Удалён ${acc.email}`, "ok");
      } catch (err) {
        toast(err.message, "err");
      }
    }
  }

  function renderMessages() {
    const container = el("messagesList");
    if (!state.messages.length) {
      container.innerHTML = `
        <div class="empty-state">
          <div class="empty-title">Писем нет</div>
          <div class="muted">${
            state.accounts.length
              ? "Попробуй нажать «Обновить» или изменить фильтр."
              : "Подключи первый аккаунт Outlook слева."
          }</div>
        </div>`;
      return;
    }

    container.innerHTML = "";
    for (const msg of state.messages) {
      const color = accountColor(msg.account_email);
      const card = document.createElement("div");
      card.className = `msg-card${msg.is_read ? "" : " unread"}${
        state.selectedMessageId === msg.id ? " active" : ""
      }`;

      const senderLabel = msg.sender_name || msg.sender_email || "Неизвестный отправитель";
      const isJunk = msg.folder === "junkemail";
      card.innerHTML = `
        <div class="msg-top">
          <div class="msg-badges">
            <span class="account-pill" style="background:${color.bg};color:${color.fg};border-color:${color.border}">
              ${escapeHtml(msg.account_email)}
            </span>
            ${isJunk ? `<span class="folder-pill">Спам</span>` : ""}
          </div>
          <span class="msg-time">${escapeHtml(formatShortTime(msg.received_at))}</span>
        </div>
        <div class="msg-sender-row">
          ${msg.is_read ? "" : `<span class="unread-dot"></span>`}
          <span class="msg-sender">${escapeHtml(senderLabel)}</span>
        </div>
        <div class="msg-subject">${escapeHtml(msg.subject || "(Без темы)")}</div>
        <div class="msg-preview">${escapeHtml(msg.body_preview || "")}</div>
      `;

      card.addEventListener("click", () => openMessage(msg.id));
      container.appendChild(card);
    }
  }

  function renderReaderHeader(msg) {
    const color = accountColor(msg.account_email);
    const pill = el("readerAccountPill");
    pill.textContent = msg.account_email;
    pill.style.background = color.bg;
    pill.style.color = color.fg;
    pill.style.borderColor = color.border;

    el("readerFolderPill").classList.toggle("hidden", msg.folder !== "junkemail");
    el("readerSubject").textContent = msg.subject || "(Без темы)";
    el("readerFrom").textContent = msg.sender_name
      ? `${msg.sender_name} <${msg.sender_email}>`
      : msg.sender_email || "—";
    el("readerTo").textContent = msg.to_recipients || msg.account_email;
    el("readerDate").textContent = formatFullDate(msg.received_at);

    const webBtn = el("btnOpenWebLink");
    if (msg.web_link) {
      webBtn.href = msg.web_link;
      webBtn.classList.remove("hidden");
    } else {
      webBtn.classList.add("hidden");
    }
  }

  function renderReaderBody(msg) {
    el("readerLoading").classList.add("hidden");
    const iframe = el("readerIframe");
    const pre = el("readerText");
    const toggleBtn = el("btnToggleBodyMode");

    const isHtml = msg.body_content_type === "html";
    toggleBtn.classList.toggle("hidden", !isHtml);

    if (isHtml && state.bodyViewMode === "html") {
      toggleBtn.textContent = "Текст";
      pre.classList.add("hidden");
      iframe.classList.remove("hidden");
      iframe.srcdoc = msg.body_content || "";
    } else {
      toggleBtn.textContent = "HTML";
      iframe.classList.add("hidden");
      pre.classList.remove("hidden");
      pre.textContent = isHtml
        ? htmlToPlainText(msg.body_content || "")
        : msg.body_content || msg.body_preview || "(Пустое письмо)";
    }
  }

  function selectView(view, accountId = null, customTitle = "") {
    state.selectedView = view;
    state.selectedAccountId = accountId;

    document.querySelectorAll(".nav-item").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.view === view);
    });

    const isVault = view === "vault";
    el("listPane").classList.toggle("hidden", isVault);
    el("readerPane").classList.toggle("hidden", isVault);
    el("vaultPane").classList.toggle("hidden", !isVault);

    if (isVault) {
      renderAccounts();
      loadVault();
      return;
    }

    const titles = {
      "all-inbox": "Входящие (все аккаунты)",
      "all-unread": "Непрочитанные (все аккаунты)",
      "all-junk": "Спам / Junk (все аккаунты)",
      "all-everything": "Все письма (все аккаунты)",
    };
    el("currentViewTitle").textContent =
      view === "account" ? `Ящик: ${customTitle}` : titles[view] || "Письма";

    renderAccounts();
    loadMessages();
  }

  function setSyncSpinner(active) {
    el("syncSpinner").classList.toggle("spinning", active);
    el("btnSyncAll").disabled = active;
    if (active) {
      el("syncStatusText").textContent = "Синхронизация...";
    }
  }

  async function triggerSyncAll() {
    if (!state.accounts.length) {
      toast("Сначала добавь хотя бы один аккаунт", "err");
      return;
    }
    setSyncSpinner(true);
    try {
      const data = await api("/api/sync", { method: "POST" });
      state.accounts = data.accounts || state.accounts;
      state.stats = data.stats || state.stats;
      state.lastGlobalSync = data.last_global_sync || state.lastGlobalSync;
      renderStats();
      renderAccounts();
      await loadMessages();

      const results = data.results || [];
      const failed = results.filter((r) => !r.ok);
      if (failed.length) {
        toast(`Обновлено с ошибками (${failed.length} акк.)`, "err");
      } else if (results.length) {
        toast(`Синхронизировано ${results.length} акк.`, "ok");
      }
    } catch (err) {
      toast(err.message, "err");
    } finally {
      setSyncSpinner(false);
    }
  }

  // --- Modals & OAuth flow ---

  function openModal(id) {
    el(id).classList.remove("hidden");
  }

  function closeModal(id) {
    el(id).classList.add("hidden");
    if (id === "modalAddAccount") {
      stopOAuthPolling();
    }
  }

  function stopOAuthPolling() {
    if (state.oauthPollTimer) {
      clearInterval(state.oauthPollTimer);
      state.oauthPollTimer = null;
    }
  }

  function setOAuthStatus(kind, text) {
    const box = el("oauthStatusBox");
    box.className = `oauth-status ${kind}`;
    box.textContent = text;
    box.classList.remove("hidden");
  }

  function openAddAccountModal(prefillEmail = "") {
    el("oauthLoginHint").value = prefillEmail || "";
    el("importEmail").value = prefillEmail || "";
    el("oauthStatusBox").classList.add("hidden");
    el("deviceStatusBox").classList.add("hidden");
    el("deviceCodeBox").classList.add("hidden");
    el("importStatusBox").classList.add("hidden");
    if (!state.settings?.configured) {
      setOAuthStatus(
        "error",
        "Для входа через Microsoft сначала вставь Azure Application (client) ID в «Настройки Azure» (или перейди во вкладку «Импорт refresh_token», если токен уже есть)."
      );
    }
    openModal("modalAddAccount");
  }

  function pollOAuthState(oauthState, setStatusFn, onSuccess = null) {
    stopOAuthPolling();
    state.oauthPollTimer = setInterval(async () => {
      try {
        const st = await api(`/api/oauth/status/${encodeURIComponent(oauthState)}`);
        if (st.status === "completed") {
          stopOAuthPolling();
          setStatusFn(
            "ok",
            `Готово! Аккаунт ${st.email} подключён, refresh_token сохранён.`
          );
          if (onSuccess) onSuccess(st);
          toast(`Подключён ${st.email}`, "ok");
          await loadAccounts();
          await loadVault();
          setTimeout(() => {
            loadAccounts();
            loadMessages();
          }, 2000);
        } else if (st.status === "error") {
          stopOAuthPolling();
          setStatusFn("error", `Ошибка: ${st.error}`);
        } else if (st.status === "expired") {
          stopOAuthPolling();
          setStatusFn("error", "Сессия авторизации истекла — попробуй снова.");
        }
      } catch {
        // ignore transient poll error
      }
    }, 1500);
  }

  async function startOAuth(mode) {
    stopOAuthPolling();
    const loginHint = el("oauthLoginHint").value.trim();
    try {
      const data = await api("/api/oauth/start", {
        method: "POST",
        body: JSON.stringify({ login_hint: loginHint, prompt: "login" }),
      });

      if (mode === "copy") {
        await copyText(
          data.authorize_url,
          "Ссылка скопирована — открой её в окне Инкогнито"
        );
        setOAuthStatus(
          "pending",
          "Жду завершения входа в Инкогнито... Как только войдёшь, аккаунт появится здесь автоматически."
        );
      } else {
        window.open(data.authorize_url, "outlook_oauth", "width=560,height=720");
        setOAuthStatus(
          "pending",
          "Окно входа Microsoft открыто. Войди в аккаунт и нажми «Принять»..."
        );
      }

      pollOAuthState(data.state, setOAuthStatus, () => {
        el("oauthLoginHint").value = "";
      });
    } catch (err) {
      setOAuthStatus("error", err.message);
    }
  }

  function setDeviceStatus(kind, text) {
    const box = el("deviceStatusBox");
    box.className = `oauth-status ${kind}`;
    box.textContent = text;
    box.classList.remove("hidden");
  }

  async function startDeviceCodeFlow() {
    stopOAuthPolling();
    el("deviceCodeBox").classList.add("hidden");
    setDeviceStatus("pending", "Запрашиваю код устройства у Microsoft...");
    try {
      const data = await api("/api/oauth/device/start", {
        method: "POST",
        body: JSON.stringify({ protocol: "graph" }),
      });
      el("deviceUserCode").textContent = data.user_code;
      el("linkDeviceVerify").href = data.verification_uri || "https://microsoft.com/devicelogin";
      el("deviceCodeBox").classList.remove("hidden");
      await copyText(data.user_code, `Код ${data.user_code} скопирован`);
      setDeviceStatus(
        "pending",
        "Открой microsoft.com/devicelogin (можно в Инкогнито или с телефона), вставь код и войди в Outlook..."
      );
      pollOAuthState(data.state, setDeviceStatus);
    } catch (err) {
      setDeviceStatus("error", err.message);
    }
  }

  async function submitManualImport() {
    const raw = el("importRefreshToken").value.trim();
    const cid = el("importClientId").value.trim();
    const email = el("importEmail").value.trim();
    const box = el("importStatusBox");
    if (!raw) {
      box.className = "oauth-status error";
      box.textContent = "Вставь refresh_token или список строк.";
      box.classList.remove("hidden");
      return;
    }
    box.className = "oauth-status pending";
    box.textContent = "Проверяю refresh_token (Graph / IMAP XOAUTH2) и загружаю письма...";
    box.classList.remove("hidden");

    const isBulk = raw.includes("\n") || (raw.includes(":") && raw.includes("@"));
    try {
      if (isBulk) {
        const res = await api("/api/accounts/import-bulk", {
          method: "POST",
          body: JSON.stringify({ raw_text: raw, protocol: "auto" }),
        });
        const errHint = res.error_count
          ? ` (ошибок: ${res.error_count} — ${res.errors[0]?.error || ""})`
          : "";
        box.className = res.imported_count > 0 ? "oauth-status ok" : "oauth-status error";
        box.textContent = `Импортировано: ${res.imported_count}${errHint}`;
        if (res.imported_count > 0) {
          el("importRefreshToken").value = "";
          toast(`Импортировано аккаунтов: ${res.imported_count}`, "ok");
          await loadAccounts();
          await loadVault();
          await loadMessages();
        }
      } else {
        const res = await api("/api/accounts/import", {
          method: "POST",
          body: JSON.stringify({
            refresh_token: raw,
            client_id: cid,
            email,
            protocol: "auto",
          }),
        });
        box.className = "oauth-status ok";
        box.textContent = `Аккаунт ${res.account.email} (${res.account.protocol || "graph"}) добавлен! Загружено писем: ${res.sync.fetched}.`;
        el("importRefreshToken").value = "";
        el("importEmail").value = "";
        toast(`Добавлен ${res.account.email}`, "ok");
        await loadAccounts();
        await loadVault();
        await loadMessages();
      }
    } catch (err) {
      box.className = "oauth-status error";
      box.textContent = `Ошибка: ${err.message}`;
    }
  }

  async function saveSettings(customPayload = null) {
    const payload = customPayload || {
      client_id: el("setClientId").value.trim(),
      tenant: el("setTenant").value.trim() || "common",
      client_secret: el("setClientSecret").value.trim() || null,
      poll_interval_sec: Number(el("setPollInterval").value) || 60,
      messages_per_folder: Number(el("setMessagesPerFolder").value) || 30,
      auto_sync: el("setAutoSync").checked,
      sync_junk: el("setSyncJunk").checked,
    };
    try {
      await api("/api/settings", {
        method: "PUT",
        body: JSON.stringify(payload),
      });
      await loadSettings();
      closeModal("modalSettings");
      toast("Настройки сохранены", "ok");
    } catch (err) {
      toast(err.message, "err");
    }
  }

  // --- Wire up events ---

  function bindEvents() {
    document.querySelectorAll(".nav-item").forEach((btn) => {
      btn.addEventListener("click", () => selectView(btn.dataset.view));
    });

    el("btnAddAccount").addEventListener("click", () => openAddAccountModal());
    el("btnOpenSettings").addEventListener("click", async () => {
      await loadSettings();
      openModal("modalSettings");
    });
    el("btnSyncAll").addEventListener("click", triggerSyncAll);

    el("btnCopyRedirectOnboarding").addEventListener("click", () =>
      copyText(state.settings?.redirect_uri || "http://localhost:8765/callback", "Redirect URI скопирован")
    );
    el("btnCopyRedirectVault").addEventListener("click", () =>
      copyText(state.settings?.redirect_uri || "http://localhost:8765/callback", "Redirect URI скопирован")
    );
    el("btnCopyRedirectSettings").addEventListener("click", () =>
      copyText(state.settings?.redirect_uri || "http://localhost:8765/callback", "Redirect URI скопирован")
    );

    el("btnSaveQuickClientId").addEventListener("click", async () => {
      const cid = el("quickClientId").value.trim();
      if (!cid) {
        toast("Вставь Application (client) ID", "err");
        return;
      }
      await saveSettings({ client_id: cid });
    });

    el("btnSaveVaultClientId").addEventListener("click", async () => {
      const cid = el("vaultQuickClientId").value.trim();
      if (!cid) {
        toast("Вставь Application (client) ID", "err");
        return;
      }
      await saveSettings({ client_id: cid });
    });

    el("btnSaveSettings").addEventListener("click", () => saveSettings());
    el("btnStartOAuthPopup").addEventListener("click", () => startOAuth("popup"));
    el("btnCopyOAuthLink").addEventListener("click", () => startOAuth("copy"));
    el("btnStartDeviceCode").addEventListener("click", startDeviceCodeFlow);
    el("btnCopyDeviceCode").addEventListener("click", () => {
      const code = el("deviceUserCode").textContent.trim();
      if (code && code !== "----") copyText(code, `Код ${code} скопирован`);
    });
    el("btnSubmitImport").addEventListener("click", submitManualImport);

    el("btnOpenVaultImport").addEventListener("click", () => {
      el("vaultImportStatus").classList.add("hidden");
      openModal("modalImportVault");
    });

    el("btnToggleVaultSecrets").addEventListener("click", () => {
      state.showVaultSecrets = !state.showVaultSecrets;
      el("btnToggleVaultSecrets").textContent = state.showVaultSecrets
        ? "Скрыть пароли"
        : "Показать пароли";
      renderVault();
    });

    el("vaultSearchInput").addEventListener("input", (e) => {
      state.vaultSearch = e.target.value;
      renderVault();
    });

    el("btnAutoLoginAllVault")?.addEventListener("click", async () => {
      const unconn = state.vaultEntries.filter((e) => !e.oauth_connected);
      if (!unconn.length) {
        toast("Все аккаунты из базы уже подключены!", "ok");
        return;
      }
      if (!confirm(`Запустить авто-вход для ${unconn.length} неподключённых аккаунтов?`)) {
        return;
      }
      const btn = el("btnAutoLoginAllVault");
      btn.disabled = true;
      const origText = btn.textContent;
      btn.textContent = "⏳ Вход...";
      toast(`Запущен авто-вход для ${unconn.length} аккаунтов...`, "ok");
      try {
        const res = await api("/api/vault/auto-login-all", { method: "POST" });
        const succ = (res.results || []).filter((r) => r.status === "ok").length;
        const fail = (res.results || []).filter((r) => r.status === "error").length;
        toast(`Авто-вход завершён: успешно ${succ}, ошибок ${fail}`, succ ? "ok" : "err");
        await loadAccounts();
        await loadVault();
        await loadStats();
      } catch (err) {
        toast(`Ошибка авто-входа: ${err.message || err}`, "err");
      } finally {
        btn.disabled = false;
        btn.textContent = origText;
      }
    });

    el("btnSubmitVaultImport").addEventListener("click", async () => {
      const rawText = el("vaultImportTextarea").value.trim();
      const autoLogin = Boolean(el("chkVaultAutoLogin")?.checked);
      const box = el("vaultImportStatus");
      if (!rawText) {
        box.className = "oauth-status error";
        box.textContent = "Вставь хотя бы одну строку с аккаунтом.";
        box.classList.remove("hidden");
        return;
      }
      try {
        const res = await api("/api/vault/import", {
          method: "POST",
          body: JSON.stringify({ raw_text: rawText, auto_login: autoLogin }),
        });
        state.vaultEntries = res.entries || [];
        state.stats = res.stats || state.stats;
        renderStats();
        renderVault();
        el("vaultImportTextarea").value = "";
        closeModal("modalImportVault");
        if (autoLogin) {
          toast(`Импортировано ${res.imported_count} акк. Фоновый авто-вход запущен!`, "ok");
          setTimeout(async () => {
            await loadAccounts();
            await loadVault();
            await loadStats();
          }, 8000);
          setTimeout(async () => {
            await loadAccounts();
            await loadVault();
            await loadStats();
          }, 20000);
        } else {
          toast(`Добавлено в базу: ${res.imported_count} акк.`, "ok");
        }
      } catch (err) {
        box.className = "oauth-status error";
        box.textContent = err.message;
        box.classList.remove("hidden");
      }
    });

    el("btnReloadRecoveryModal").addEventListener("click", () => {
      const entry = state.vaultEntries.find((e) => e.id === state.activeRecoveryEntryId);
      if (!entry) return;
      const customHost = el("recoveryImapHostInput").value.trim();
      fetchRecoveryMail(entry, customHost, true);
    });

    el("btnToggleBodyMode").addEventListener("click", () => {
      if (!state.activeMessage) return;
      state.bodyViewMode = state.bodyViewMode === "html" ? "text" : "html";
      renderReaderBody(state.activeMessage);
    });
    el("btnReloadBody").addEventListener("click", () => {
      if (state.selectedMessageId) {
        openMessage(state.selectedMessageId, true);
      }
    });

    let searchTimer = null;
    el("searchInput").addEventListener("input", (e) => {
      state.searchQuery = e.target.value;
      clearTimeout(searchTimer);
      searchTimer = setTimeout(loadMessages, 200);
    });

    el("accountFilterInput").addEventListener("input", (e) => {
      state.accountFilter = e.target.value;
      renderAccounts();
    });

    document.querySelectorAll("[data-close-modal]").forEach((btn) => {
      btn.addEventListener("click", () => closeModal(btn.dataset.closeModal));
    });
    document.querySelectorAll(".modal-backdrop").forEach((backdrop) => {
      backdrop.addEventListener("click", (e) => {
        if (e.target === backdrop) closeModal(backdrop.id);
      });
    });

    document.querySelectorAll(".tab").forEach((tab) => {
      tab.addEventListener("click", () => {
        const target = tab.dataset.tab;
        document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
        tab.classList.add("active");
        document.querySelectorAll(".tab-panel").forEach((p) => p.classList.add("hidden"));
        el(target).classList.remove("hidden");
      });
    });

    window.addEventListener("message", (event) => {
      if (event.origin !== window.location.origin) return;
      if (event.data?.type === "outlook-oauth-done" && event.data.ok) {
        loadAccounts();
        loadVault();
        setTimeout(loadMessages, 1800);
      }
    });
  }

  async function init() {
    bindEvents();
    await loadSettings();
    await loadAccounts();
    await loadVault();

    if (state.accounts.length === 0 && state.vaultEntries.length > 0) {
      selectView("vault");
    } else {
      await loadMessages();
    }

    // Live 1s countdown + auto-refresh TOTP codes on 30-second boundary
    state.totpTickTimer = setInterval(() => {
      if (document.hidden || state.selectedView !== "vault" || !state.vaultEntries.length) {
        return;
      }
      const nowSec = Math.floor(Date.now() / 1000);
      const step = Math.floor(nowSec / 30);
      const remaining = 30 - (nowSec % 30);
      document.querySelectorAll("[data-totp-timer]").forEach((node) => {
        node.textContent = `${remaining}с`;
      });
      if (step !== state.lastTotpStep) {
        state.lastTotpStep = step;
        loadVault().catch(() => {});
      }
    }, 1000);

    // Gently keep UI in sync with background auto-sync loop on server
    state.uiPollTimer = setInterval(async () => {
      if (document.hidden) return;
      try {
        const prevSync = state.lastGlobalSync;
        await loadAccounts();
        if (state.lastGlobalSync !== prevSync) {
          await loadMessages();
        }
      } catch {
        // ignore transient errors
      }
    }, 12000);
  }

  init();
})();
