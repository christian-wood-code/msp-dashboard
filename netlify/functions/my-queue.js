// ─── PERSONAL AUTOTASK QUEUE ────────────────────────────────────────────────
// Returns the calling user's own open tickets (grouped overdue / due today /
// rest) plus a compact roster of every other team resource with their open
// count and a deep link to their queue in the Autotask web UI.
//
// Auth: the caller sends their own Microsoft Graph access token (the same
// token the browser already holds from signing in via MSAL for
// Calendars.Read/Mail.Read/Tasks.Read/Chat.Read). We call Graph's own
// `/me` endpoint with that token server-side to learn who's really asking —
// a live call to Microsoft, not a client-supplied claim, so it can't be
// spoofed by passing someone else's email as a query param. The resulting
// email is matched against Autotask Resources.email (case-insensitive).
//
// "Open" = every status except Complete (5); Work Complete (16) still counts
// as open — same rule as daily-ticket-status's digest, confirmed with
// Christian 2026-08-10 (see that project's memory before changing this).
const at = require('./_autotask.js');

const CLOSED_STATUS = 5;
const NZ_TZ = 'Pacific/Auckland';

function nzDateKey(date) {
  return new Intl.DateTimeFormat('en-CA', { timeZone: NZ_TZ, year: 'numeric', month: '2-digit', day: '2-digit' }).format(date);
}

async function whoAmI(graphToken) {
  const res = await fetch('https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName', {
    headers: { Authorization: `Bearer ${graphToken}` }
  });
  if (!res.ok) {
    const err = new Error(`Graph /me failed: ${res.status}`);
    err.statusCode = res.status === 401 ? 401 : 502;
    throw err;
  }
  const me = await res.json();
  const email = (me.mail || me.userPrincipalName || '').toLowerCase();
  if (!email) throw Object.assign(new Error('No email on this Microsoft account'), { statusCode: 422 });
  return { email, displayName: me.displayName || email };
}

function bucketTicket(t, now, todayKey) {
  if (!t.dueDateTime) return 'rest';
  const due = new Date(t.dueDateTime);
  if (due < now) return 'overdue';
  if (nzDateKey(due) === todayKey) return 'dueToday';
  return 'rest';
}

exports.handler = async (event) => {
  const authHeader = event.headers?.authorization || event.headers?.Authorization || '';
  const graphToken = authHeader.replace(/^Bearer\s+/i, '');
  if (!graphToken) return { statusCode: 401, body: JSON.stringify({ error: 'Missing Microsoft Graph token' }) };
  if (!at.hasCredentials()) return { statusCode: 500, body: JSON.stringify({ error: 'Autotask credentials not configured' }) };

  try {
    const me = await whoAmI(graphToken);
    const { baseUrl, webUrl } = await at.getZoneInfo();

    const [tickets, resources] = await Promise.all([
      at.query(baseUrl, 'Tickets',
        [{ op: 'noteq', field: 'status', value: CLOSED_STATUS }],
        ['id', 'ticketNumber', 'title', 'dueDateTime', 'assignedResourceID']),
      at.query(baseUrl, 'Resources', [{ op: 'gte', field: 'id', value: 0 }], ['id', 'firstName', 'lastName', 'email'])
    ]);

    const resourceByEmail = new Map(resources.map(r => [(r.email || '').toLowerCase(), r]));
    const myResource = resourceByEmail.get(me.email);

    const openByResource = new Map();
    tickets.forEach(t => {
      const rid = t.assignedResourceID;
      if (!openByResource.has(rid)) openByResource.set(rid, []);
      openByResource.get(rid).push(t);
    });

    // "My queue" in the Autotask web UI isn't parameterized by resource id at
    // all — it's tied to whoever is logged into that Autotask session, e.g.
    // https://ww29.autotask.net/AutotaskOnyx/LandingPage?view=my-tasks-and-tickets&view-data=<base64 of
    // {"url":"https://ww29.autotask.net/Mvc/ServiceDesk/MyWorkspace.mvc/MyTaskAndTicket"}>
    // (confirmed against a real copied URL, 2026-09). Same URL for every
    // viewer, works because each person is signed into their own Autotask
    // session in their own browser. There is deliberately no per-resource
    // equivalent below — Autotask exposes no URL to open *someone else's*
    // queue, so team rows get a count only, no broken deep link.
    const myWorkspaceUrl =
      `${webUrl}/AutotaskOnyx/LandingPage?view=my-tasks-and-tickets&view-data=` +
      encodeURIComponent(Buffer.from(JSON.stringify({ url: `${webUrl}/Mvc/ServiceDesk/MyWorkspace.mvc/MyTaskAndTicket` })).toString('base64'));

    const now = new Date();
    const todayKey = nzDateKey(now);
    const mine = myResource ? (openByResource.get(myResource.id) || []) : [];
    const grouped = { overdue: [], dueToday: [], rest: [] };
    mine.forEach(t => grouped[bucketTicket(t, now, todayKey)].push({
      id: t.id, title: t.title || `Ticket #${t.ticketNumber}`,
      dueDateTime: t.dueDateTime || null, url: at.ticketUrl(webUrl, t.id)
    }));

    const team = resources
      .filter(r => r.id !== myResource?.id)
      .map(r => ({
        name: (r.firstName || '').trim() || `Resource #${r.id}`,
        openCount: (openByResource.get(r.id) || []).length
      }))
      .filter(r => r.openCount > 0)
      .sort((a, b) => b.openCount - a.openCount);

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        me: { name: me.displayName, matched: Boolean(myResource) },
        mine: myResource ? {
          queueUrl: myWorkspaceUrl,
          openCount: mine.length,
          overdue: grouped.overdue,
          dueToday: grouped.dueToday,
          rest: grouped.rest
        } : null,
        team
      })
    };
  } catch (err) {
    return { statusCode: err.statusCode || 500, body: JSON.stringify({ error: err.message }) };
  }
};
