// Hộp thư Outlook đa tài khoản — gọi API backend (Microsoft Graph + OAuth)
const API = "/api/outlook";

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

function escapeHtml(s) {
  return (s || "").replace(/[&<>"]/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

// ---------- Tài khoản ----------
async function loadStatus() {
  const s = await api("/status");
  el("not-configured").style.display = s.configured ? "none" : "block";
  renderAccounts(s.accounts || []);
  return s;
}

function renderAccounts(accounts) {
  const box = el("accounts");
  if (!accounts.length) {
    box.innerHTML = `<p class="hint">Chưa có tài khoản nào. Bấm "+ Thêm tài khoản".</p>`;
    return;
  }
  box.innerHTML = accounts.map((a) => `
    <div class="acct-chip">
      <span>✉️</span>
      <span class="email">${escapeHtml(a.username)}</span>
      <button class="danger small" data-id="${a.home_account_id}">Xóa</button>
    </div>`).join("");
  box.querySelectorAll("button.danger").forEach((b) => {
    b.onclick = async () => {
      if (!confirm("Xóa tài khoản này? (sẽ phải đăng nhập lại nếu thêm sau này)")) return;
      await api("/accounts/" + encodeURIComponent(b.dataset.id), { method: "DELETE" });
      loadStatus();
    };
  });
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
    setTimeout(() => pollLogin(jobId), 2500); // pending -> poll tiếp
  } catch (e) {
    el("login-status").textContent = "❌ " + e.message;
    el("btn-add").disabled = false;
  }
}

// ---------- Hộp thư / Tìm kiếm ----------
async function loadInbox() {
  const box = el("inbox");
  const unread = el("chk-unread").checked;
  box.innerHTML = `<p class="hint">Đang tải mail…</p>`;
  try {
    const data = await api(`/inbox?count=20&unread_only=${unread}`);
    if (!data.inboxes.length) {
      box.innerHTML = `<p class="hint">Chưa có tài khoản nào. Hãy thêm tài khoản trước.</p>`;
      return;
    }
    renderGroups(data.inboxes);
  } catch (e) {
    box.innerHTML = `<p class="hint">❌ ${escapeHtml(e.message)}</p>`;
  }
}

async function doSearch() {
  const q = el("search-input").value.trim();
  if (!q) { loadInbox(); return; }
  const box = el("inbox");
  box.innerHTML = `<p class="hint">Đang tìm "${escapeHtml(q)}"…</p>`;
  el("btn-clear").style.display = "inline-block";
  try {
    const data = await api(`/search?q=${encodeURIComponent(q)}&count=20`);
    if (!data.results.length) {
      box.innerHTML = `<p class="hint">Không có tài khoản nào.</p>`;
      return;
    }
    renderGroups(data.results);
  } catch (e) {
    box.innerHTML = `<p class="hint">❌ ${escapeHtml(e.message)}</p>`;
  }
}

function clearSearch() {
  el("search-input").value = "";
  el("btn-clear").style.display = "none";
  loadInbox();
}

function renderGroups(groups) {
  el("inbox").innerHTML = groups.map(renderGroup).join("");
  // Gắn sự kiện click mở chi tiết
  el("inbox").querySelectorAll(".mail-item").forEach((node) => {
    node.onclick = () => openMail(node.dataset.acct, node.dataset.id);
  });
}

function renderGroup(group) {
  const head = `<div class="group-head">📥 ${escapeHtml(group.account)}</div>`;
  if (group.error) return head + `<p class="hint">❌ ${escapeHtml(group.error)}</p>`;
  if (!group.messages.length) return head + `<p class="hint">Không có mail.</p>`;
  const items = group.messages.map((m) => `
    <div class="mail-item ${m.is_read ? "" : "unread"}"
         data-acct="${group.home_account_id}" data-id="${m.id}">
      <p class="mail-subject">${escapeHtml(m.subject)}</p>
      <p class="mail-meta">${escapeHtml(m.from_name || m.from_address)} · ${fmtDate(m.received)}</p>
      <p class="mail-preview">${escapeHtml(m.preview)}</p>
    </div>`).join("");
  return head + items;
}

// ---------- Chi tiết mail (modal) ----------
let currentMail = null; // { acct, id, is_read }

async function openMail(acct, id) {
  el("modal").style.display = "flex";
  el("m-subject").textContent = "Đang tải…";
  el("m-meta").textContent = "";
  el("m-body").srcdoc = "";
  try {
    const m = await api(`/message/${encodeURIComponent(acct)}/${encodeURIComponent(id)}`);
    currentMail = { acct, id, is_read: m.is_read };
    el("m-subject").textContent = m.subject;
    const to = m.to && m.to.length ? ` → ${m.to.map(escapeHtml).join(", ")}` : "";
    el("m-meta").innerHTML =
      `${escapeHtml(m.from_name || m.from_address)}${to} · ${fmtDate(m.received)}`;
    el("m-weblink").href = m.web_link || "#";
    // Render body trong iframe sandbox (không chạy script -> an toàn)
    el("m-body").srcdoc = m.body_type === "html"
      ? m.body
      : `<pre style="white-space:pre-wrap;font-family:system-ui;padding:12px;">${escapeHtml(m.body)}</pre>`;
    updateToggleLabel();
    // Mở mail thì coi như đã đọc
    if (!m.is_read) setRead(true, false);
  } catch (e) {
    el("m-subject").textContent = "Lỗi";
    el("m-body").srcdoc = `<p style="padding:12px;">${escapeHtml(e.message)}</p>`;
  }
}

function updateToggleLabel() {
  el("m-toggle-read").textContent = currentMail?.is_read
    ? "Đánh dấu chưa đọc" : "Đánh dấu đã đọc";
}

async function setRead(isRead, refresh = true) {
  const mail = currentMail; // chụp tham chiếu, tránh bị null hóa giữa chừng
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
    alert("Không đổi được trạng thái: " + e.message);
  }
}

function refreshCurrentView() {
  const q = el("search-input").value.trim();
  if (q) doSearch(); else loadInbox();
}

function closeModal() {
  el("modal").style.display = "none";
  currentMail = null;
  refreshCurrentView();
}

// ---------- Khởi tạo ----------
el("btn-add").onclick = startLogin;
el("btn-refresh").onclick = loadInbox;
el("chk-unread").onchange = loadInbox;
el("btn-search").onclick = doSearch;
el("btn-clear").onclick = clearSearch;
el("search-input").addEventListener("keydown", (e) => { if (e.key === "Enter") doSearch(); });
el("m-close").onclick = closeModal;
el("m-toggle-read").onclick = () => { if (currentMail) setRead(!currentMail.is_read); };
el("modal").addEventListener("click", (e) => { if (e.target.id === "modal") closeModal(); });

loadStatus();
