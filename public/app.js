/* ─── JARVIS DASHBOARD — per-viewer, real Microsoft sign-in ────────────────
   Replaces the old claude.ai-Artifact/connector version: this is a normal
   website, so every viewer signs in with their own Microsoft account (MSAL,
   auth-code + PKCE, no client secret) and the browser calls Microsoft Graph
   directly for their calendar/inbox/tasks/chats. The one server call
   (netlify/functions/my-queue.js) forwards that same Graph access token to
   verify identity and returns the caller's own Autotask ticket queue plus a
   compact team roster — Autotask credentials never reach the browser.
*/
window.addEventListener("error", e => console.error("WINDOW ERROR:", e.message, "@", e.filename + ":" + e.lineno));
window.addEventListener("unhandledrejection", e => console.error("UNHANDLED REJECTION:", e.reason?.stack || e.reason?.message || String(e.reason)));
const boom = (where, e) => console.error("SECTION FAIL [" + where + "]:", e?.stack || e?.message || String(e));

const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const el = id => document.getElementById(id);
const GRAPH = "https://graph.microsoft.com/v1.0";

/* ---------- auth ---------- */
const msalApp = new msal.PublicClientApplication({
  auth: {
    clientId: window.JARVIS_CONFIG.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${window.JARVIS_CONFIG.TENANT_ID}`,
    redirectUri: window.JARVIS_CONFIG.REDIRECT_URI
  },
  cache: { cacheLocation: "sessionStorage" }
});

let account = null;

async function ensureSignedIn() {
  await msalApp.initialize();
  const result = await msalApp.handleRedirectPromise();
  if (result?.account) account = result.account;
  else {
    const existing = msalApp.getAllAccounts();
    if (existing.length) account = existing[0];
  }
  if (!account) {
    el("bootMsg").textContent = "SIGN-IN REQUIRED";
    el("signInBtn").hidden = false;
    el("signInBtn").addEventListener("click", () => msalApp.loginRedirect({ scopes: window.JARVIS_CONFIG.SCOPES }));
    return false;
  }
  return true;
}

async function getToken() {
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: window.JARVIS_CONFIG.SCOPES, account });
    return r.accessToken;
  } catch (e) {
    // Silent refresh needs an interactive step (expired session, new scope, etc).
    await msalApp.acquireTokenRedirect({ scopes: window.JARVIS_CONFIG.SCOPES, account });
    return null; // page is navigating away
  }
}

async function graph(path) {
  const token = await getToken();
  if (!token) return null;
  const res = await fetch(GRAPH + path, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Graph ${path} failed: ${res.status} ${body.slice(0, 200)}`);
  }
  return res.json();
}

async function myQueue() {
  const token = await getToken();
  if (!token) return null;
  const res = await fetch("/.netlify/functions/my-queue", { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body.error || `my-queue failed: ${res.status}`);
  }
  return res.json();
}

/* ---------- time ---------- */
const fmtTime = d => d.toLocaleTimeString("en-NZ", { hour: "numeric", minute: "2-digit", hour12: true }).replace(/\s*(am|pm|AM|PM)$/, m => m.trim().toLowerCase());
const fmtDay = d => d.toLocaleDateString([], { weekday: "short", day: "numeric", month: "short" });
const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
function ago(d) {
  const h = (Date.now() - d.getTime()) / 36e5;
  if (h < 1) return Math.max(1, Math.round(h * 60)) + "m ago";
  if (h < 24) return Math.round(h) + "h ago";
  const dy = Math.round(h / 24); return dy + (dy === 1 ? " day ago" : " days ago");
}
// Graph timestamps in calendarView are in the requested Prefer:outlook.timezone
// (we ask for Pacific/Auckland below) but arrive without a zone suffix, so
// treat them as wall-clock NZ time rather than re-interpreting as UTC.
function toDate(dt) {
  if (!dt) return null;
  // Graph's calendarView returns start/end.dateTime in UTC with NO "Z" or
  // offset suffix (e.g. "2026-09-03T01:00:00.0000000") since no
  // Prefer:outlook.timezone header is sent. Without appending "Z" here,
  // `new Date(...)` parses that string as *local* time instead of UTC --
  // a 1pm NZT meeting (1am UTC) then displays as "1:00am", a 12-hour flip.
  // Confirmed against a real mailbox, 2026-09.
  let s = dt.replace(" ", "T").replace(/(\.\d{3})\d+$/, "$1"); // trim Graph's 7-digit fractional seconds — extra digits make Date parsing fail
  if (!/[zZ]|[+-]\d{2}:\d{2}$/.test(s)) s += "Z";
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

/* ---------- core readout ---------- */
(function tick() {
  const now = new Date(), h = now.getHours();
  const greet = h < 12 ? "Good morning" : h < 18 ? "Good afternoon" : "Good evening";
  el("core").innerHTML =
    '<div class="cloc">AUCKLAND NZ</div>' +
    '<div class="ctime">' + esc(now.toLocaleTimeString("en-NZ", { hour: "numeric", minute: "2-digit", second: "2-digit", hour12: true }).replace(/\s*(am|pm|AM|PM)$/, m => " " + m.trim().toLowerCase())) + '</div>' +
    '<div class="cdate">' + esc(now.toLocaleDateString([], { weekday: "short", day: "numeric", month: "short" })) + '</div>' +
    '<div class="cgreet">' + esc(greet + (account ? "," : "")) + '</div>' +
    (account ? '<div class="cname">' + esc((account.name || "").split(" ")[0]) + '</div>' : '');
  setTimeout(tick, 1000);
})();

/* ---------- sections ---------- */
const R = 34; // dial radius, % of container
function hexPoint(k) { // k = 0..5, 0 = top, clockwise
  const a = (k * 60) * Math.PI / 180;
  return { x: +(50 + R * Math.sin(a)).toFixed(1), y: +(50 - R * Math.cos(a)).toFixed(1) };
}
const SECTIONS = [
  { key: "today",    label: "Today",      glyph: "◷", c: "#35e0f5", ...hexPoint(0) },
  { key: "tasks",    label: "To Do",      glyph: "▣", c: "#ffd24a", ...hexPoint(1) },
  { key: "queue",    label: "Queue",      glyph: "▤", c: "#ff8a3d", ...hexPoint(2) },
  { key: "upcoming", label: "Coming Up",  glyph: "◈", c: "#b98cff", ...hexPoint(3) },
  { key: "inbox",    label: "Inbox",      glyph: "✉", c: "#ff5b5b", ...hexPoint(4) },
  { key: "teams",    label: "Chats",      glyph: "◉", c: "#33e08a", ...hexPoint(5) }
];
const DATA = { today: [], upcoming: [], mail: [], todo: [], chats: [], queue: null, err: {} };
// Always opens on Today, deliberately not remembered across visits — a new
// sign-in (or reload) should show what's happening now, not wherever the
// dial was last left.
let ACTIVE = "today";

const dial = el("dial");
for (const s of SECTIONS) {
  const b = document.createElement("button");
  b.className = "lobe"; b.type = "button"; b.dataset.key = s.key;
  b.style.setProperty("--c", s.c);
  b.style.left = s.x + "%"; b.style.top = s.y + "%";
  b.setAttribute("aria-selected", String(s.key === ACTIVE));
  b.innerHTML = `<span class="lglyph">${s.glyph}</span><span class="lcount" data-count>–</span><span class="llabel">${esc(s.label)}</span>`;
  b.addEventListener("click", () => select(s.key));
  dial.appendChild(b);
}
document.addEventListener("keydown", e => {
  if (e.key !== "ArrowLeft" && e.key !== "ArrowRight") return;
  const i = SECTIONS.findIndex(s => s.key === ACTIVE);
  const n = (i + (e.key === "ArrowRight" ? 1 : SECTIONS.length - 1)) % SECTIONS.length;
  select(SECTIONS[n].key);
});
function select(key) {
  ACTIVE = key;
  dial.querySelectorAll(".lobe").forEach(b => b.setAttribute("aria-selected", String(b.dataset.key === key)));
  paint();
}
function setCount(key, v) {
  const b = dial.querySelector(`.lobe[data-key="${key}"] [data-count]`);
  if (b) b.textContent = v;
}

/* ---------- row builders ---------- */
function eventRows(list, withDay) {
  const now = new Date();
  if (!list.length) return `<div class="empty">${withDay ? "Nothing scheduled ahead." : "Diary clear today."}</div>`;
  return list.map(e => {
    const s = toDate(e.start?.dateTime), en = toDate(e.end?.dateTime);
    const live = s && en && now >= s && now <= en;
    const joinUrl = e.onlineMeeting?.joinUrl || e.onlineMeetingUrl || null;
    const others = (e.attendees || []).map(a => a.emailAddress?.name).filter(Boolean);
    const who = others.length ? others.slice(0, 3).join(", ") + (others.length > 3 ? ` +${others.length - 3}` : "") : "solo";
    return `<div class="row ${e.isCancelled ? "cancelled" : ""}">
      <div class="tick ${live ? "now" : ""}"></div>
      <div class="time">${s ? esc(fmtTime(s)) : "—"}${withDay && s ? `<small>${esc(fmtDay(s))}</small>` : en ? `<small>to ${esc(fmtTime(en))}</small>` : ""}</div>
      <div class="rmain">
        <div class="rtitle">${live ? '<span class="pill new">LIVE</span>' : ""}<a href="${esc(e.webLink || "#")}" target="_blank" rel="noopener">${esc(e.subject || "(no subject)")}</a></div>
        <div class="rmeta">${e.isCancelled ? "CANCELLED · " : ""}${esc(who)}${joinUrl ? ` · <a class="pill join" href="${esc(joinUrl)}" target="_blank" rel="noopener">JOIN</a>` : ""}</div>
      </div></div>`;
  }).join("");
}
function mailRows() {
  if (!DATA.mail.length) return '<div class="empty">Nothing new here.</div>';
  return DATA.mail.map(m => {
    const d = m.receivedDateTime ? new Date(m.receivedDateTime) : null;
    return `<div class="row">
      <div class="tick ${m.importance === "high" ? "act" : ""}"></div>
      <div class="rmain">
        <div class="rtitle">${m.importance === "high" ? '<span class="pill act">! IMPORTANT</span>' : ""}${!m.isRead ? '<span class="pill new">NEW</span>' : ""}<a href="${esc(m.webLink || "#")}" target="_blank" rel="noopener">${esc(m.subject || "(no subject)")}</a></div>
        <div class="rmeta">${esc(m.from?.emailAddress?.name || "?")} · ${d ? esc(ago(d)) : ""}</div>
      </div></div>`;
  }).join("");
}
function todoRows() {
  if (!DATA.todo.length) return '<div class="empty">Nothing open in Microsoft To Do.</div>';
  const todayISO = new Date().toISOString().slice(0, 10);
  const sorted = DATA.todo.slice().sort((a, b) => (a.dueDateTime?.dateTime || "9999").localeCompare(b.dueDateTime?.dateTime || "9999"));
  return sorted.map(t => {
    const due = t.dueDateTime?.dateTime ? t.dueDateTime.dateTime.slice(0, 10) : null;
    const overdue = due && due < todayISO;
    const important = t.importance === "high";
    return `<div class="row">
      <div class="tick ${overdue ? "stale" : important ? "act" : ""}"></div>
      <div class="rmain">
        <div class="rtitle" style="white-space:normal">${esc(t.title)}
          ${important ? '<span class="pill act" style="margin-left:7px">★ HIGH</span>' : ""}
          ${due ? `<span class="pill age ${overdue ? "bad" : "warn"}">${overdue ? "OVERDUE " : ""}${esc(new Date(due).toLocaleDateString([], { day: "numeric", month: "short" }))}</span>` : ""}</div>
      </div></div>`;
  }).join("");
}
function chatRows() {
  if (!DATA.chats.length) return '<div class="empty">No recent chat activity.</div>';
  const me = (account?.username || "").toLowerCase();
  return DATA.chats.map(c => {
    // Channel name for a named group chat, otherwise the other person/people's
    // name(s) — excluding yourself by comparing each member's email against
    // your own signed-in account (Graph chat members have no "own" flag).
    const others = (c.members || []).filter(m => (m.email || "").toLowerCase() !== me);
    const title = c.topic || others.map(m => m.displayName).filter(Boolean).join(", ") || "Chat";
    const preview = c.lastMessagePreview;
    const when = preview?.createdDateTime ? new Date(preview.createdDateTime) : null;
    const snip = (preview?.body?.content || "").replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim().slice(0, 64);
    const href = `https://teams.microsoft.com/l/chat/${encodeURIComponent(c.id)}/conversations`;
    return `<div class="row">
      <div class="tick"></div>
      <div class="rmain">
        <div class="rtitle"><a href="${esc(href)}" target="_blank" rel="noopener">${esc(title)}</a></div>
        <div class="rmeta">${when ? esc(ago(when)) + " · " : ""}${esc(snip)}</div>
      </div></div>`;
  }).join("");
}
function queueRows() {
  const q = DATA.queue;
  if (!q) return '<div class="empty">Loading your queue…</div>';
  if (!q.me?.matched || !q.mine) {
    return `<div class="empty">No Autotask resource matches your signed-in email${q.me?.name ? " (" + esc(q.me.name) + ")" : ""}. Ask Christian to check the Resources.email field in Autotask.</div>`;
  }
  const mine = q.mine;
  const section = (label, cls, items) => !items.length ? "" : `<div class="divider">${label} · ${items.length}</div>` +
    items.map(t => `<div class="row"><div class="tick ${cls}"></div><div class="rmain">
        <div class="rtitle"><a href="${esc(t.url)}" target="_blank" rel="noopener">${esc(t.title)}</a></div>
        ${t.dueDateTime ? `<div class="rmeta">Due ${esc(new Date(t.dueDateTime).toLocaleDateString([], { day: "numeric", month: "short" }))}</div>` : ""}
      </div></div>`).join("");

  const mineHtml = mine.openCount === 0
    ? `<div class="row"><div class="tick ok"></div><div class="rmain"><div class="rtitle">Queue clear <span class="pill ok">0 OPEN</span></div></div></div>`
    : section("Overdue", "stale", mine.overdue) + section("Due today", "act", mine.dueToday) + section("Open", "", mine.rest);

  // Autotask exposes no URL to open someone else's queue (its "my queue" view
  // is tied to whichever session is logged in, not a resource id) — so
  // teammate rows show a count only, deliberately no link.
  const teamHtml = !q.team.length ? "" : '<div class="divider">Team queues</div>' + q.team.map(r => `
    <div class="row">
      <div class="tick ${r.openCount > 10 ? "stale" : r.openCount > 4 ? "act" : ""}"></div>
      <div class="counts"><div class="stat"><b>${r.openCount}</b><span>Open</span></div></div>
      <div class="rmain"><div class="rtitle">${esc(r.name)}</div></div>
    </div>`).join("");

  return `<a class="open-link" style="margin:12px 15px;display:inline-block" href="${esc(mine.queueUrl)}" target="_blank" rel="noopener">OPEN MY QUEUE IN AUTOTASK ▸</a>` + mineHtml + teamHtml;
}

/* ---------- paint ---------- */
let LOADING = true;
function paint() {
  const s = SECTIONS.find(x => x.key === ACTIVE);
  el("panel").style.setProperty("--acc", s.c);
  el("pGlyph").textContent = s.glyph;
  el("pTitle").textContent = s.label;
  el("pTabs").innerHTML = "";

  const err = DATA.err[ACTIVE];
  let html, count = "–";
  if (err) { html = `<div class="err">UPLINK FAULT · ${esc(err)}</div>`; count = "!"; }
  else if (LOADING && ACTIVE !== "queue") html = '<div class="skel">Scanning</div>';
  else if (ACTIVE === "today") { html = eventRows(DATA.today, false); count = DATA.today.filter(e => !e.isCancelled).length; }
  else if (ACTIVE === "upcoming") { html = eventRows(DATA.upcoming, true); count = DATA.upcoming.length; }
  else if (ACTIVE === "inbox") { html = mailRows(); count = DATA.mail.filter(m => !m.isRead).length; }
  else if (ACTIVE === "tasks") { html = todoRows(); count = DATA.todo.length; }
  else if (ACTIVE === "teams") { html = chatRows(); count = DATA.chats.length; }
  else if (ACTIVE === "queue") { html = queueRows(); count = DATA.queue?.mine?.openCount ?? (DATA.queue ? "0" : "…"); }
  el("pBody").innerHTML = html;
  el("pCount").textContent = count;
}

/* ---------- assessment ---------- */
function buildAssessment() {
  const now = new Date();
  const bits = [];
  const withSpan = e => ({ e, s: toDate(e.start?.dateTime), en: toDate(e.end?.dateTime) });
  const liveNow = DATA.today.filter(e => !e.isCancelled).map(withSpan).find(x => x.s && x.en && now >= x.s && now <= x.en);
  const next = DATA.today.filter(e => !e.isCancelled).map(withSpan).filter(x => x.s && x.s > now).sort((a, b) => a.s - b.s)[0];
  if (liveNow) bits.push(`You're in ${liveNow.e.subject || "a meeting"} until ${fmtTime(liveNow.en)}.`);
  else if (next) bits.push(`Next on the diary: ${next.e.subject || "an appointment"} at ${fmtTime(next.s)}.`);
  else bits.push("Nothing left on today's diary.");

  const unread = DATA.mail.filter(m => !m.isRead).length;
  if (unread) bits.push(`${unread} unread email${unread === 1 ? "" : "s"}.`);

  const overdueTickets = DATA.queue?.mine?.overdue?.length || 0;
  if (overdueTickets) bits.push(`${overdueTickets} Autotask ticket${overdueTickets === 1 ? " is" : "s are"} overdue.`);
  else if (DATA.queue?.mine) bits.push("Autotask queue clear of overdue tickets.");

  return bits.join(" ");
}

/* ---------- load ---------- */
async function load() {
  paint();

  const now = new Date();
  const start = new Date(now); start.setHours(0, 0, 0, 0);
  const end = new Date(start); end.setDate(end.getDate() + 8);

  const pCal = graph(`/me/calendarView?startDateTime=${start.toISOString()}&endDateTime=${end.toISOString()}&$orderby=start/dateTime&$top=50&$select=subject,start,end,isCancelled,attendees,webLink,onlineMeeting`)
    .then(r => {
      if (!r) return;
      const events = r.value || [];
      // "Today" is a morning-only readout: 12am-12pm. A same-day event after
      // noon isn't dropped, just moved into "Coming Up" instead, so nothing
      // on the real calendar disappears — it just isn't called "today".
      const isThisMorning = d => sameDay(d, now) && d.getHours() < 12;
      DATA.today = events.filter(e => { const d = toDate(e.start?.dateTime); return d && isThisMorning(d); });
      DATA.upcoming = events.filter(e => { const d = toDate(e.start?.dateTime); return d && d > now && !isThisMorning(d); }).slice(0, 10);
      setCount("today", DATA.today.filter(e => !e.isCancelled).length);
      setCount("upcoming", DATA.upcoming.length);
    })
    .catch(e => { boom("calendar", e); DATA.err.today = DATA.err.upcoming = e.message; setCount("today", "!"); setCount("upcoming", "!"); });

  const pMail = graph(`/me/mailFolders/Inbox/messages?$top=25&$orderby=receivedDateTime desc&$select=subject,from,receivedDateTime,isRead,webLink,importance`)
    .then(r => {
      if (!r) return;
      // Important-flagged mail first, then newest-first within each group —
      // Graph's own $orderby can't express "important first" so this is
      // done client-side after the (already newest-first) fetch.
      DATA.mail = (r.value || []).slice().sort((a, b) =>
        (b.importance === "high") - (a.importance === "high") ||
        new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
      setCount("inbox", DATA.mail.filter(m => !m.isRead).length);
    })
    .catch(e => { boom("inbox", e); DATA.err.inbox = e.message; setCount("inbox", "!"); });

  // Fetching /me/todo/lists/{id}/tasks with the list id spliced into the URL
  // path fails outright — confirmed 2026-09 against a real mailbox, 400
  // "RequestBroker--ParseUri" even with the id correctly percent-encoded and
  // every query param removed. Some Outlook/To Do resource ids trip up
  // Graph's gateway when placed in a path segment. Fixed by never putting
  // the id in a path at all: $expand=tasks pulls each list's tasks inline
  // via a single /me/todo/lists request instead.
  const pTodo = graph(`/me/todo/lists?$select=id&$expand=tasks($select=title,dueDateTime,importance,status)`)
    .then(r => {
      if (!r) return;
      const all = (r.value || []).flatMap(list => list.tasks || []);
      DATA.todo = all.filter(t => t.status !== "completed");
      setCount("tasks", DATA.todo.length);
    })
    .catch(e => { boom("tasks", e); DATA.err.tasks = e.message; setCount("tasks", "!"); });

  const pChats = graph(`/me/chats?$top=15&$select=id,topic,chatType&$expand=members,lastMessagePreview&$orderby=lastMessagePreview/createdDateTime desc`)
    .then(r => { if (!r) return; DATA.chats = (r.value || []).filter(c => c.lastMessagePreview); setCount("teams", DATA.chats.length); })
    .catch(e => { boom("teams", e); DATA.err.teams = e.message; setCount("teams", "!"); });

  const pQueue = myQueue()
    .then(q => { if (!q) return; DATA.queue = q; setCount("queue", q.mine?.openCount ?? 0); if (ACTIVE === "queue") paint(); })
    .catch(e => { boom("queue", e); DATA.err.queue = e.message; setCount("queue", "!"); });

  await Promise.all([pCal, pMail, pTodo, pChats, pQueue]);
  LOADING = false;
  paint();

  el("assessText").className = "";
  try { el("assessText").textContent = buildAssessment(); }
  catch (e) { el("assessText").textContent = "Summary unavailable — the panels are current."; boom("assessment", e); }
}

/* ---------- boot ---------- */
(async function start() {
  const signedIn = await ensureSignedIn();
  if (!signedIn) return; // waiting on the sign-in button / redirect
  el("boot").classList.add("hide");
  el("wrap").hidden = false;
  el("wrap").classList.add("rise");
  el("signOutBtn").addEventListener("click", () => msalApp.logoutRedirect());
  load();
})();
