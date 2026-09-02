// ─── SHARED AUTOTASK REST API CLIENT ──────────────────────────────────────────
// All field names, entity names, auth headers and query format verified against
// the official Autotask REST API v1.0 documentation.
//
// Auth: three headers on every request — UserName, Secret, ApiIntegrationCode.
// Zone lookup returns a base URL that ALREADY includes "/ATServicesRest",
// so entity paths are appended as "/v1.0/EntityName/...".

// Uses the global fetch built into Node 18+ (no node-fetch dependency).

const USERNAME         = process.env.AUTOTASK_USERNAME;
const SECRET           = process.env.AUTOTASK_SECRET;
const INTEGRATION_CODE = process.env.AUTOTASK_INTEGRATION_CODE || '';

function hasCredentials() {
  return Boolean(USERNAME && SECRET && INTEGRATION_CODE);
}

function atHeaders() {
  return {
    'UserName':           USERNAME,
    'Secret':             SECRET,
    'ApiIntegrationCode': INTEGRATION_CODE,
    'Content-Type':       'application/json',
    'Accept':             'application/json'
  };
}

// First call in any integration: resolve the tenant's zone-specific base URL.
// zoneInformation also returns `webUrl` (the human-facing Autotask portal
// domain, e.g. "ww29.autotask.net") alongside the API domain
// ("webservices29.autotask.net") -- the zone numbers happen to match but
// don't assume that; use the value the API actually gives you.
async function getZoneInfo() {
  const res = await fetch(
    `https://webservices.autotask.net/atservicesrest/v1.0/zoneInformation?user=${encodeURIComponent(USERNAME)}`
  );
  if (!res.ok) throw new Error(`Zone lookup failed: ${res.status} ${res.statusText}`);
  const data = await res.json();
  return {
    baseUrl: data.url.replace(/\/$/, ''),    // e.g. "https://webservices29.autotask.net/ATServicesRest"
    webUrl: data.webUrl.replace(/\/$/, '')   // e.g. "https://ww29.autotask.net"
  };
}

async function getZoneUrl() {
  return (await getZoneInfo()).baseUrl;
}

// Direct link to a ticket's detail page in the Autotask web UI.
// Format confirmed via Autotask docs (ExecuteCommand / dynamic variables):
// Mvc/ServiceDesk/TicketDetail.mvc?workspace=False&mode=0&TicketID={id}
// Uses the numeric ticket `id`, NOT the human-facing `ticketNumber`.
function ticketUrl(webUrl, ticketId) {
  return `${webUrl}/Mvc/ServiceDesk/TicketDetail.mvc?workspace=False&mode=0&TicketID=${ticketId}`;
}

// Query an entity. First page via POST, subsequent pages via GET nextPageUrl.
// `extra` may set { maxRecords } for capped queries.
async function query(baseUrl, entity, filter = [], fields = [], extra = {}) {
  const search = {
    filter: filter.length ? filter : [{ op: 'gte', field: 'id', value: 0 }]
  };
  if (fields.length) search.includeFields = fields;
  if (extra.maxRecords) search.maxRecords = extra.maxRecords;

  const firstRes = await fetch(`${baseUrl}/v1.0/${entity}/query`, {
    method: 'POST',
    headers: atHeaders(),
    body: JSON.stringify(search)
  });
  if (!firstRes.ok) {
    const body = await firstRes.text().catch(() => '');
    throw new Error(`${entity} query failed: ${firstRes.status} ${body.slice(0, 300)}`);
  }
  const firstJson = await firstRes.json();
  let results = firstJson.items || [];

  let nextUrl = firstJson.pageDetails?.nextPageUrl || null;
  while (nextUrl && !extra.maxRecords) {
    const res = await fetch(nextUrl, { headers: atHeaders() });
    if (!res.ok) break;
    const json = await res.json();
    results = results.concat(json.items || []);
    nextUrl = json.pageDetails?.nextPageUrl || null;
  }
  return results;
}

// Like query() but returns [] on error instead of throwing (for optional endpoints).
async function querySafe(baseUrl, entity, filter = [], fields = [], extra = {}) {
  return query(baseUrl, entity, filter, fields, extra).catch(err => {
    console.warn(`[${entity}] skipped: ${err.message}`);
    return [];
  });
}

// Fetch an entity's field definitions (incl. picklist label/value pairs).
async function fieldInfo(baseUrl, entity) {
  const res = await fetch(`${baseUrl}/v1.0/${entity}/entityInformation/fields`, {
    headers: atHeaders()
  });
  if (!res.ok) return [];
  const json = await res.json();
  return json.fields || [];
}

// Build a { value: label } map from a picklist field.
function picklistMap(fields, fieldName) {
  const f = fields.find(x => x.name === fieldName);
  const map = {};
  if (f && Array.isArray(f.picklistValues)) {
    f.picklistValues.forEach(p => { map[p.value] = p.label; });
  }
  return map;
}

module.exports = {
  hasCredentials, atHeaders, getZoneUrl, getZoneInfo, ticketUrl,
  query, querySafe, fieldInfo, picklistMap
};
