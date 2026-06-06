// Hộp thư Outlook đa tài khoản — gọi API backend (Microsoft Graph + OAuth)
const API = "/api/outlook";
const PAGE = 20;

const el = (id) => document.getElementById(id);

async function api(path, opts) {
  const res = await fetch(API + path, opts);
  if (!res.ok) {
    let detail = res.statusText;
    try { detail = (await res.json()).detail || detail; } catch (_) {}
    throw new Error(detail);
  }
  return res.json();
}

function fmtDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  return d.toLocaleString("vi-VN", { dateStyle: "short", timeStyle: "short" });
}

function fmtSize(n) {
  if (!n) return "";
  if (n < 1024) return n + " B";
  if (n < 1048576) return (n / 1024).toFixed(0) + " KB";
  return (n / 1048576).toFixed(1) + " MB";
}

function escapeHtml(s) {
  return (s || "").replace(/[&<>"]/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

// ---------- State ----------
const state = {
  folder: "inbox",
  searchMode: false,
  order: [],        // [{accId, name}]
  loaded: {},       // accId -> [messages]
  pages: {},        // accId -> số trang đã tải (để tính $skip server)
  hasMore: {},      // accId -> bool
  accounts: [],
  tokenUpdateAccountId: null,
  accountHealth: {},
  importJob: {
    id: null,
    status: "",
    offset: 0,
    limit: 50,
    poll: null,
    lines: [],
    allResults: [],
  },
};

function filterParams() {
  const p = new URLSearchParams();
  p.set("folder", state.folder);
  p.set("count", PAGE);
  if (el("f-unread").checked) p.set("unread_only", "true");
  if (el("f-att").checked) p.set("has_attachments", "true");
  const sender = el("f-sender").value.trim();
  if (sender) p.set("sender", sender);
  if (el("f-from").value) p.set("date_from", el("f-from").value);
  if (el("f-to").value) p.set("date_to", el("f-to").value);
  const acc = el("f-account").value;
  if (acc) p.set("account", acc);
  return p;
}

// ---------- Tài khoản ----------
async function loadStatus() {
  const s = await api("/status");
  el("not-configured").style.display = s.configured ? "none" : "block";
  state.accounts = s.accounts || [];
  renderAccounts(s.accounts || []);
  if (s.accounts && s.accounts.length) loadUnreadCounts();
  return s;
}

function renderAccounts(accounts) {
  const box = el("accounts");
  // Cập nhật dropdown lọc theo tài khoản
  const sel = el("f-account");
  const cur = sel.value;
  sel.innerHTML = `<option value="">Tất cả</option>` +
    accounts.map((a) => `<option value="${a.home_account_id}">${escapeHtml(a.username)}</option>`).join("");
  sel.value = cur;

  if (!accounts.length) {
    box.innerHTML = `<p class="hint">Chưa có tài khoản nào. Bấm "+ Thêm tài khoản".</p>`;
    return;
  }
  box.innerHTML = `
    <div class="account-table-wrap">
      <table class="account-table">
        <thead>
          <tr>
            <th>Email</th>
            <th>Nguồn</th>
            <th>Scope</th>
            <th>Cập nhật</th>
            <th>Chưa đọc</th>
            <th>Health</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          ${accounts.map((a) => accountRowHtml(a)).join("")}
        </tbody>
      </table>
    </div>`;
  box.querySelectorAll("button.danger").forEach((b) => {
    b.onclick = async () => {
      if (!confirm("Xóa tài khoản này? (sẽ phải đăng nhập lại nếu thêm sau này)")) return;
      await api("/accounts/" + encodeURIComponent(b.dataset.id), { method: "DELETE" });
      loadStatus();
    };
  });
  box.querySelectorAll("[data-update-token]").forEach((b) => {
    b.onclick = () => openTokenModal(b.dataset.updateToken);
  });
  box.querySelectorAll("[data-check-account]").forEach((b) => {
    b.onclick = () => checkAccountHealth(b.dataset.checkAccount);
  });
}

function accountRowHtml(a) {
  const id = escapeHtml(a.home_account_id);
  const source = a.source ? `<span class="source">${escapeHtml(a.source)}</span>` : "";
  const scope = a.scope ? `<span class="scope">${escapeHtml(shortScope(a.scope))}</span>` : "—";
  const updated = a.updated_at ? fmtDate(a.updated_at) : "—";
  const health = state.accountHealth[a.home_account_id];
  const healthHtml = health
    ? `<span class="health ${health.ok ? "ok" : "fail"}" title="${escapeHtml(health.detail || "")}">${escapeHtml(health.label || health.status)}</span>`
    : `<span class="health unknown">—</span>`;
  const updateButton = a.can_update_token
    ? `<button class="btn small" data-update-token="${id}">Đổi token</button>`
    : `<button class="btn small" disabled>Đổi token</button>`;
  return `
    <tr>
      <td class="account-email">${escapeHtml(a.username)}</td>
      <td>${source}</td>
      <td>${scope}</td>
      <td>${escapeHtml(updated)}</td>
      <td><span class="badge zero" data-badge="${id}"></span></td>
      <td>${healthHtml}</td>
      <td class="account-actions">
        <button class="btn small" data-check-account="${id}">Kiểm tra</button>
        ${updateButton}
        <button class="danger small" data-id="${id}">Xóa</button>
      </td>
    </tr>`;
}

function shortScope(scope) {
  return (scope || "")
    .replaceAll("https://graph.microsoft.com/", "")
    .replaceAll("https://outlook.office.com/", "");
}

async function checkAccountHealth(accountId) {
  state.accountHealth[accountId] = {
    ok: false,
    status: "checking",
    label: "Đang kiểm tra",
    detail: "",
  };
  renderAccounts(state.accounts);
  try {
    const result = await api(`/accounts/${encodeURIComponent(accountId)}/health`);
    state.accountHealth[accountId] = result;
  } catch (e) {
    state.accountHealth[accountId] = {
      ok: false,
      status: "error",
      label: "Lỗi",
      detail: e.message,
    };
  }
  renderAccounts(state.accounts);
  loadUnreadCounts();
}

async function importRefreshToken() {
  const input = el("token-line");
  const lines = input.value.split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
  startImportLines(lines);
}

function getImportConcurrency() {
  const raw = Number.parseInt(el("import-concurrency").value, 10);
  if (Number.isNaN(raw)) return 3;
  return Math.min(Math.max(raw, 1), 5);
}

async function startImportLines(lines) {
  const status = el("token-import-status");
  const button = el("btn-import-token");
  if (!lines.length) {
    status.textContent = "Chưa có dòng token.";
    return;
  }
  const badLine = lines.findIndex((line) => line.split("|").length !== 4);
  if (badLine >= 0) {
    status.textContent = `Dòng ${badLine + 1} sai định dạng email|password|refresh_token|client_id.`;
    return;
  }
  button.disabled = true;
  stopImportPolling();
  state.importJob.id = null;
  state.importJob.status = "";
  state.importJob.offset = 0;
  state.importJob.lines = lines;
  state.importJob.allResults = [];
  renderImportLog([]);
  renderImportPager(null);
  renderImportProgress(null);
  updateFailedActions();
  status.textContent = `Đang tạo job import ${lines.length} token…`;
  try {
    const job = await api("/accounts/import-refresh-token/job", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        line: lines.join("\n"),
        concurrency: getImportConcurrency(),
      }),
    });
    state.importJob.id = job.job_id;
    state.importJob.status = job.status;
    el("btn-cancel-import").style.display = "inline-block";
    renderImportJob(job);
    state.importJob.poll = setInterval(() => pollImportJob(), 1500);
    pollImportJob();
  } catch (e) {
    status.textContent = "Lỗi: " + e.message;
    renderImportLog([]);
    renderImportPager(null);
    renderImportProgress(null);
    button.disabled = false;
    el("btn-cancel-import").style.display = "none";
  }
}

function stopImportPolling() {
  if (state.importJob.poll) {
    clearInterval(state.importJob.poll);
    state.importJob.poll = null;
  }
}

async function pollImportJob() {
  if (!state.importJob.id) return;
  try {
    const p = new URLSearchParams();
    p.set("offset", state.importJob.offset);
    p.set("limit", state.importJob.limit);
    const job = await api(`/accounts/import-jobs/${state.importJob.id}?${p.toString()}`);
    renderImportJob(job);
    if (["done", "cancelled"].includes(job.status)) {
      stopImportPolling();
      el("btn-import-token").disabled = false;
      el("btn-cancel-import").style.display = "none";
      if (job.failed === 0 && job.status === "done") el("token-line").value = "";
      await loadAllImportResults(job);
      updateFailedActions();
      await loadStatus();
    }
  } catch (e) {
    stopImportPolling();
    el("token-import-status").textContent = "Lỗi job import: " + e.message;
    el("btn-import-token").disabled = false;
    el("btn-cancel-import").style.display = "none";
  }
}

function renderImportJob(job) {
  state.importJob.status = job.status;
  state.importJob.offset = job.offset || 0;
  const active = ["queued", "running", "cancelling"].includes(job.status);
  const doneText = job.status === "done"
    ? "Hoàn tất"
    : job.status === "cancelled"
      ? "Đã hủy"
      : job.status === "cancelling"
        ? "Đang hủy"
        : "Đang chạy";
  el("token-import-status").textContent =
    `${doneText}: ${job.processed}/${job.total}. Thêm: ${job.added}. Lỗi: ${job.failed}. Luồng: ${job.concurrency || 1}.`;
  renderImportProgress(job);
  renderImportLog(job.results || []);
  renderImportPager(job);
  el("btn-import-token").disabled = active;
  el("btn-cancel-import").style.display = active ? "inline-block" : "none";
}

async function loadAllImportResults(job) {
  if (!job || !job.job_id || !job.result_count) {
    state.importJob.allResults = [];
    return;
  }
  const all = [];
  for (let offset = 0; offset < job.result_count; offset += 200) {
    const p = new URLSearchParams();
    p.set("offset", offset);
    p.set("limit", 200);
    const page = await api(`/accounts/import-jobs/${job.job_id}?${p.toString()}`);
    all.push(...(page.results || []));
  }
  state.importJob.allResults = all.sort((a, b) => a.line - b.line);
}

function failedImportLines() {
  return (state.importJob.allResults || [])
    .filter((item) => !item.ok)
    .map((item) => state.importJob.lines[item.line - 1])
    .filter(Boolean);
}

function updateFailedActions() {
  const failed = failedImportLines();
  const show = failed.length > 0 && !["queued", "running", "cancelling"].includes(state.importJob.status);
  el("btn-copy-failed").style.display = show ? "inline-block" : "none";
  el("btn-retry-failed").style.display = show ? "inline-block" : "none";
}

async function copyFailedLines() {
  const failed = failedImportLines();
  if (!failed.length) return;
  const text = failed.join("\n");
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
    } else {
      const ta = document.createElement("textarea");
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      ta.remove();
    }
    el("token-import-status").textContent = `Đã copy ${failed.length} dòng lỗi.`;
  } catch (e) {
    el("token-import-status").textContent = "Không copy được: " + e.message;
  }
}

function retryFailedLines() {
  const failed = failedImportLines();
  if (!failed.length) return;
  el("token-line").value = failed.join("\n");
  startImportLines(failed);
}

async function checkEmailsLive() {
  const input = el("check-emails-input");
  const status = el("check-emails-status");
  const button = el("btn-check-emails");
  const lines = input.value.split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
  if (!lines.length) {
    status.textContent = "Chưa có email để kiểm tra.";
    renderEmailCheckResults([]);
    return;
  }
  button.disabled = true;
  status.textContent = `Đang kiểm tra ${lines.length} email…`;
  try {
    const data = await api("/check-emails", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ emails: lines.join("\n") }),
    });
    status.textContent = `Kết quả: ${data.ok}/${data.total} domain có khả năng nhận mail.`;
    renderEmailCheckResults(data.results || []);
  } catch (e) {
    status.textContent = "Lỗi: " + e.message;
    renderEmailCheckResults([]);
  } finally {
    button.disabled = false;
  }
}

function renderEmailCheckResults(results) {
  const box = el("check-emails-results");
  if (!results || !results.length) {
    box.style.display = "none";
    box.innerHTML = "";
    return;
  }
  box.style.display = "block";
  box.innerHTML = `
    <div class="check-table-wrap">
      <table class="check-table">
        <thead>
          <tr>
            <th>Dòng</th>
            <th>Email</th>
            <th>Trạng thái</th>
            <th>Provider</th>
            <th>MX</th>
            <th>Ghi chú</th>
          </tr>
        </thead>
        <tbody>
          ${results.map((item) => `
            <tr class="${item.ok ? "ok" : "fail"}">
              <td>${item.line}</td>
              <td>${escapeHtml(item.email || "")}</td>
              <td>${emailCheckStatusText(item)}</td>
              <td>${escapeHtml(item.provider || "unknown")}</td>
              <td>${escapeHtml((item.mx || []).slice(0, 3).join(", "))}</td>
              <td>${escapeHtml(item.note || "")}</td>
            </tr>`).join("")}
        </tbody>
      </table>
    </div>`;
}

function emailCheckStatusText(item) {
  if (item.status === "domain_can_receive") return "Có thể nhận";
  if (item.status === "syntax_error") return "Sai format";
  if (item.status === "nxdomain") return "Domain không tồn tại";
  if (item.status === "no_mx") return "Không có MX";
  return "Không rõ";
}

function renderImportProgress(job) {
  const box = el("import-progress");
  if (!job) {
    box.style.display = "none";
    box.innerHTML = "";
    return;
  }
  const pct = job.total ? Math.round((job.processed / job.total) * 100) : 0;
  box.style.display = "block";
  box.innerHTML = `
    <div class="progress-row">
      <div class="progress-bar"><span style="width:${pct}%"></span></div>
      <span>${pct}%</span>
    </div>`;
}

function renderImportPager(job) {
  const box = el("import-pager");
  if (!job || !job.result_count) {
    box.style.display = "none";
    box.innerHTML = "";
    return;
  }
  const start = job.offset + 1;
  const end = Math.min(job.offset + job.limit, job.result_count);
  const prevDisabled = job.offset <= 0 ? "disabled" : "";
  const nextDisabled = end >= job.result_count ? "disabled" : "";
  box.style.display = "flex";
  box.innerHTML = `
    <button class="btn small" id="import-prev" ${prevDisabled}>Trước</button>
    <span class="hint">Dòng kết quả ${start}-${end} / ${job.result_count}</span>
    <button class="btn small" id="import-next" ${nextDisabled}>Sau</button>`;
  el("import-prev").onclick = () => {
    state.importJob.offset = Math.max(0, state.importJob.offset - state.importJob.limit);
    pollImportJob();
  };
  el("import-next").onclick = () => {
    state.importJob.offset += state.importJob.limit;
    pollImportJob();
  };
}

async function cancelImportJob() {
  if (!state.importJob.id) return;
  el("btn-cancel-import").disabled = true;
  el("token-import-status").textContent = "Đang gửi yêu cầu hủy…";
  try {
    const job = await api(`/accounts/import-jobs/${state.importJob.id}/cancel`, {
      method: "POST",
    });
    renderImportJob(job);
  } catch (e) {
    el("token-import-status").textContent = "Không hủy được: " + e.message;
  } finally {
    el("btn-cancel-import").disabled = false;
  }
}

function renderImportLog(results) {
  const box = el("import-log");
  if (!results || !results.length) {
    box.style.display = "none";
    box.innerHTML = "";
    return;
  }
  const details = document.querySelector(".token-import");
  if (details) details.open = true;
  box.style.display = "block";
  box.innerHTML = `
    <div class="import-log-table">
      <table>
        <thead>
          <tr>
            <th>Dòng</th>
            <th>Email</th>
            <th>Kết quả</th>
            <th>Chi tiết</th>
          </tr>
        </thead>
        <tbody>
          ${results.map((item) => `
            <tr class="${item.ok ? "ok" : "fail"}">
              <td>${item.line}</td>
              <td>${escapeHtml(item.email || item.account?.username || "")}</td>
              <td>${item.ok ? "OK" : "Lỗi"}</td>
              <td>${item.ok
                ? `${escapeHtml(item.source || item.account?.source || "")} ${escapeHtml(shortScope(item.scope || item.account?.scope || ""))}`
                : escapeHtml(item.error || "")}</td>
            </tr>`).join("")}
        </tbody>
      </table>
    </div>`;
}

function openTokenModal(accountId) {
  const account = state.accounts.find((a) => a.home_account_id === accountId);
  if (!account || !account.can_update_token) return;
  state.tokenUpdateAccountId = accountId;
  el("token-account").textContent = `${account.username} · ${account.source || ""} · ${shortScope(account.scope || "")}`;
  el("token-update-value").value = "";
  el("token-update-status").textContent = "";
  el("token-modal").style.display = "flex";
  setTimeout(() => el("token-update-value").focus(), 0);
}

function closeTokenModal() {
  el("token-modal").style.display = "none";
  state.tokenUpdateAccountId = null;
}

async function saveTokenUpdate() {
  const accountId = state.tokenUpdateAccountId;
  const value = el("token-update-value").value.trim();
  const status = el("token-update-status");
  const button = el("token-save");
  if (!accountId) return;
  if (!value) {
    status.textContent = "Chưa nhập refresh_token mới.";
    return;
  }
  button.disabled = true;
  status.textContent = "Đang kiểm tra token mới…";
  try {
    const data = await api(`/accounts/${encodeURIComponent(accountId)}/refresh-token`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ refresh_token: value }),
    });
    const account = data.account || {};
    status.textContent = `Đã cập nhật ${account.username || "tài khoản"} (${account.source || "token"} ${shortScope(account.scope || "")}).`;
    renderImportLog([{
      line: 1,
      email: account.username,
      ok: true,
      source: account.source,
      scope: account.scope,
      account,
    }]);
    await loadStatus();
    setTimeout(closeTokenModal, 900);
  } catch (e) {
    status.textContent = "Lỗi: " + e.message;
    renderImportLog([{
      line: 1,
      email: state.accounts.find((a) => a.home_account_id === accountId)?.username || "",
      ok: false,
      error: e.message,
    }]);
  } finally {
    button.disabled = false;
  }
}

async function loadUnreadCounts() {
  try {
    const data = await api("/unread-counts");
    data.counts.forEach((c) => {
      const badge = document.querySelector(`[data-badge="${c.home_account_id}"]`);
      if (!badge) return;
      if (c.unread == null) { badge.textContent = ""; return; }
      badge.textContent = c.unread;
      badge.classList.toggle("zero", c.unread === 0);
    });
  } catch (_) { /* badge là phụ, lỗi thì bỏ qua */ }
}

// ---------- Đăng nhập (device code flow) ----------
async function startLogin() {
  el("btn-add").disabled = true;
  try {
    const flow = await api("/login/start", { method: "POST" });
    el("login-area").style.display = "block";
    el("login-link").href = flow.verification_uri;
    el("login-code").textContent = flow.user_code;
    el("login-status").textContent = "Đang chờ bạn hoàn tất đăng nhập…";
    pollLogin(flow.job_id);
  } catch (e) {
    alert("Lỗi: " + e.message);
    el("btn-add").disabled = false;
  }
}

async function pollLogin(jobId) {
  try {
    const st = await api("/login/status/" + jobId);
    if (st.status === "success") {
      el("login-status").textContent = "✅ Đã thêm: " + (st.account?.username || "");
      setTimeout(() => { el("login-area").style.display = "none"; }, 2500);
      el("btn-add").disabled = false;
      loadStatus();
      return;
    }
    if (st.status === "error") {
      el("login-status").textContent = "❌ " + (st.error || "Đăng nhập thất bại.");
      el("btn-add").disabled = false;
      return;
    }
    setTimeout(() => pollLogin(jobId), 2500);
  } catch (e) {
    el("login-status").textContent = "❌ " + e.message;
    el("btn-add").disabled = false;
  }
}

// ---------- Hộp thư ----------
async function loadInbox() {
  state.searchMode = false;
  el("btn-clear").style.display = "none";
  el("search-input").value = "";
  const box = el("inbox");
  box.innerHTML = `<p class="hint">Đang tải mail…</p>`;
  try {
    const p = filterParams();
    p.set("skip", 0);
    const data = await api("/inbox?" + p.toString());
    if (!data.inboxes.length) {
      box.innerHTML = `<p class="hint">Chưa có tài khoản nào. Hãy thêm tài khoản trước.</p>`;
      return;
    }
    state.order = [];
    state.loaded = {}; state.pages = {}; state.hasMore = {};
    data.inboxes.forEach((g) => {
      state.order.push({ accId: g.home_account_id, name: g.account, error: g.error });
      state.loaded[g.home_account_id] = g.messages || [];
      state.pages[g.home_account_id] = 1;
      state.hasMore[g.home_account_id] = g.has_more;
    });
    renderInbox();
  } catch (e) {
    box.innerHTML = `<p class="hint">❌ ${escapeHtml(e.message)}</p>`;
  }
}

async function loadMore(accId) {
  try {
    const p = filterParams();
    p.set("account", accId);
    p.set("skip", state.pages[accId] * PAGE);
    const data = await api("/inbox?" + p.toString());
    const g = data.inboxes.find((x) => x.home_account_id === accId);
    if (g) {
      state.loaded[accId] = state.loaded[accId].concat(g.messages || []);
      state.pages[accId] += 1;
      state.hasMore[accId] = g.has_more;
    }
    renderInbox();
  } catch (e) {
    alert("Không tải thêm được: " + e.message);
  }
}

function renderInbox() {
  const box = el("inbox");
  if (!state.order.length) {
    box.innerHTML = `<p class="hint">Không có mail.</p>`;
    return;
  }
  box.innerHTML = state.order.map((g) => {
    const head = `<div class="group-head">📥 ${escapeHtml(g.name)}</div>`;
    if (g.error) return head + `<p class="hint">❌ ${escapeHtml(g.error)}</p>`;
    const msgs = state.loaded[g.accId] || [];
    if (!msgs.length) return head + `<p class="hint">Không có mail.</p>`;
    const items = msgs.map((m) => mailItemHtml(g.accId, m)).join("");
    const more = state.hasMore[g.accId]
      ? `<button class="btn small load-more" data-more="${g.accId}">⬇ Tải thêm</button>`
      : "";
    return head + items + more;
  }).join("");
  bindInboxEvents();
}

function mailItemHtml(accId, m) {
  const clip = m.has_attachments ? ` <span class="clip">📎</span>` : "";
  return `
    <div class="mail-item ${m.is_read ? "" : "unread"}" data-acct="${accId}" data-id="${m.id}">
      <p class="mail-subject">${escapeHtml(m.subject)}${clip}</p>
      <p class="mail-meta">${escapeHtml(m.from_name || m.from_address)} · ${fmtDate(m.received)}</p>
      <p class="mail-preview">${escapeHtml(m.preview)}</p>
    </div>`;
}

function bindInboxEvents() {
  el("inbox").querySelectorAll(".mail-item").forEach((node) => {
    node.onclick = () => openMail(node.dataset.acct, node.dataset.id);
  });
  el("inbox").querySelectorAll("[data-more]").forEach((b) => {
    b.onclick = () => loadMore(b.dataset.more);
  });
}

// ---------- Tìm kiếm ----------
async function doSearch() {
  const q = el("search-input").value.trim();
  if (!q) { loadInbox(); return; }
  state.searchMode = true;
  el("btn-clear").style.display = "inline-block";
  const box = el("inbox");
  box.innerHTML = `<p class="hint">Đang tìm "${escapeHtml(q)}"…</p>`;
  try {
    const data = await api(`/search?q=${encodeURIComponent(q)}&count=20`);
    if (!data.results.length) {
      box.innerHTML = `<p class="hint">Không có tài khoản nào.</p>`;
      return;
    }
    box.innerHTML = data.results.map((g) => {
      const head = `<div class="group-head">🔎 ${escapeHtml(g.account)}</div>`;
      if (g.error) return head + `<p class="hint">❌ ${escapeHtml(g.error)}</p>`;
      if (!g.messages.length) return head + `<p class="hint">Không có kết quả.</p>`;
      return head + g.messages.map((m) => mailItemHtml(g.home_account_id, m)).join("");
    }).join("");
    bindInboxEvents();
  } catch (e) {
    box.innerHTML = `<p class="hint">❌ ${escapeHtml(e.message)}</p>`;
  }
}

function clearSearch() {
  el("search-input").value = "";
  el("btn-clear").style.display = "none";
  loadInbox();
}

// ---------- Chi tiết mail (modal) ----------
let currentMail = null;

async function openMail(acct, id) {
  el("modal").style.display = "flex";
  el("m-subject").textContent = "Đang tải…";
  el("m-meta").textContent = "";
  el("m-attachments").style.display = "none";
  el("m-attachments").innerHTML = "";
  el("m-body").srcdoc = "";
  try {
    const m = await api(`/message/${encodeURIComponent(acct)}/${encodeURIComponent(id)}`);
    currentMail = { acct, id, is_read: m.is_read };
    el("m-subject").textContent = m.subject;
    const to = m.to && m.to.length ? ` → ${m.to.map(escapeHtml).join(", ")}` : "";
    el("m-meta").innerHTML = `${escapeHtml(m.from_name || m.from_address)}${to} · ${fmtDate(m.received)}`;
    el("m-weblink").href = m.web_link || "#";
    el("m-body").srcdoc = m.body_type === "html"
      ? m.body
      : `<pre style="white-space:pre-wrap;font-family:system-ui;padding:12px;">${escapeHtml(m.body)}</pre>`;
    updateToggleLabel();
    if (m.has_attachments) loadAttachments(acct, id);
    if (!m.is_read) setRead(true, false);
  } catch (e) {
    el("m-subject").textContent = "Lỗi";
    el("m-body").srcdoc = `<p style="padding:12px;">${escapeHtml(e.message)}</p>`;
  }
}

async function loadAttachments(acct, id) {
  try {
    const data = await api(`/message/${encodeURIComponent(acct)}/${encodeURIComponent(id)}/attachments`);
    const atts = (data.attachments || []).filter((a) => !a.is_inline);
    if (!atts.length) return;
    const box = el("m-attachments");
    box.style.display = "flex";
    box.innerHTML = atts.map((a) => {
      const url = `${API}/message/${encodeURIComponent(acct)}/${encodeURIComponent(id)}/attachments/${encodeURIComponent(a.id)}`;
      return `<a class="att-chip" href="${url}" download>📎 ${escapeHtml(a.name)} <span class="sz">${fmtSize(a.size)}</span></a>`;
    }).join("");
  } catch (_) { /* đính kèm là phụ */ }
}

function updateToggleLabel() {
  el("m-toggle-read").textContent = currentMail?.is_read ? "Đánh dấu chưa đọc" : "Đánh dấu đã đọc";
}

async function setRead(isRead, refresh = true) {
  const mail = currentMail;
  if (!mail) return;
  try {
    await api(
      `/message/${encodeURIComponent(mail.acct)}/${encodeURIComponent(mail.id)}/read?is_read=${isRead}`,
      { method: "POST" }
    );
    mail.is_read = isRead;
    if (currentMail === mail) updateToggleLabel();
    if (refresh) refreshCurrentView();
  } catch (e) {
    if (!refresh && e.message.includes("Mail.Read")) return;
    alert("Không đổi được trạng thái: " + e.message);
  }
}

function refreshCurrentView() {
  loadUnreadCounts();
  if (state.searchMode) doSearch(); else loadInbox();
}

function closeModal() {
  el("modal").style.display = "none";
  currentMail = null;
  refreshCurrentView();
}

// ---------- Sự kiện ----------
el("btn-add").onclick = startLogin;
el("btn-import-token").onclick = importRefreshToken;
el("btn-cancel-import").onclick = cancelImportJob;
el("btn-copy-failed").onclick = copyFailedLines;
el("btn-retry-failed").onclick = retryFailedLines;
el("btn-check-emails").onclick = checkEmailsLive;
el("btn-refresh").onclick = () => (state.searchMode ? doSearch() : loadInbox());
el("btn-search").onclick = doSearch;
el("btn-clear").onclick = clearSearch;
el("search-input").addEventListener("keydown", (e) => { if (e.key === "Enter") doSearch(); });

el("btn-filters").onclick = () => {
  const f = el("filters");
  f.style.display = f.style.display === "none" ? "block" : "none";
};
el("btn-apply").onclick = loadInbox;
el("btn-reset").onclick = () => {
  el("f-account").value = ""; el("f-sender").value = "";
  el("f-from").value = ""; el("f-to").value = "";
  el("f-unread").checked = false; el("f-att").checked = false;
  loadInbox();
};

el("folder-tabs").querySelectorAll(".ftab").forEach((b) => {
  b.onclick = () => {
    el("folder-tabs").querySelectorAll(".ftab").forEach((x) => x.classList.remove("active"));
    b.classList.add("active");
    state.folder = b.dataset.folder;
    loadInbox();
  };
});

el("m-close").onclick = closeModal;
el("m-toggle-read").onclick = () => { if (currentMail) setRead(!currentMail.is_read); };
el("modal").addEventListener("click", (e) => { if (e.target.id === "modal") closeModal(); });
el("token-close").onclick = closeTokenModal;
el("token-save").onclick = saveTokenUpdate;
el("token-modal").addEventListener("click", (e) => { if (e.target.id === "token-modal") closeTokenModal(); });

loadStatus();
