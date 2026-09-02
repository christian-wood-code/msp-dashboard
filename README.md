# Jarvis Dashboard

A per-person operational dashboard for the MSP/EUC team: each viewer signs in
with their own Microsoft account (real MSAL sign-in, not a shared connector)
to see their own calendar, inbox, Microsoft To Do tasks and Teams chats — plus
a live personal Autotask ticket queue and a compact view of teammates' open
counts.

Same visual language as the earlier Jarvis prototype, but a normal website
instead of a claude.ai Artifact — that's what makes genuinely live, per-person
Autotask data possible (an Artifact's browser sandbox can only reach a
viewer's connected claude.ai connectors, not Autotask).

## How it works

- **Microsoft 365 data** (calendar/inbox/to-do/chats): the browser calls
  Microsoft Graph directly using a token from the viewer's own sign-in
  (MSAL.js, auth-code + PKCE — no client secret, nothing server-side).
- **Autotask queue**: one Netlify function
  (`netlify/functions/my-queue.js`) takes that same Graph token, calls
  Graph's `/me` itself to verify who's really asking, matches their email to
  an Autotask `Resources` record, and returns their own tickets plus a
  roster of teammates' open counts. Autotask credentials live only in
  Netlify environment variables — never in the browser.

## One-time setup

1. **Azure AD app registration** (Entra admin center → App registrations):
   platform **Single-page application**, redirect URI = this site's deployed
   URL (add once you know it), delegated Graph scopes `User.Read`,
   `Calendars.Read`, `Mail.Read`, `Tasks.Read`, `Chat.Read` — grant admin
   consent if your tenant requires it for `Chat.Read`.
2. Paste the **Application (client) ID** and **Directory (tenant) ID** into
   `public/config.js`.
3. Autotask credentials — reuse the same API user as `autotask-report` /
   `daily-ticket-status` (`.env`, see `.env.example`).
4. Each teammate needs their Autotask `Resources.email` to match their
   Microsoft 365 sign-in email, or their queue panel won't find a match.

## Local dev

```bash
npm install
npm run dev
```

Serves the site at `http://localhost:8789` and proxies
`/.netlify/functions/*` to the same handler code Netlify runs in production.
Sign-in won't work locally until `localhost:8789/` is also added as a
redirect URI on the app registration (fine to add both the local and
production URLs).

## Deploy

Same CLI-only GitHub + Netlify pattern as the sibling projects — see
`DEPLOY.md` once written, or reuse the steps already worked out for
`daily-ticket-status`.
