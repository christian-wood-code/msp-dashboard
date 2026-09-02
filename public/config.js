// ─── AZURE AD APP REGISTRATION CONFIG ───────────────────────────────────────
// Fill these in once the "Jarvis Dashboard" app registration exists in Entra
// admin center (see the parent chat for exact steps). REDIRECT_URI must be
// added as a "Single-page application" platform redirect URI on that
// registration, matching this site's deployed URL exactly (including
// trailing slash or lack of one).
window.JARVIS_CONFIG = {
  CLIENT_ID: 'd6e9208f-537d-43fd-a509-6545dc4bdb20',
  TENANT_ID: '5f422f6e-00b5-499a-b139-dcb4042cb3ea',
  REDIRECT_URI: window.location.origin + '/',
  // Delegated Microsoft Graph scopes this dashboard needs. Must match what
  // was granted (and consented) on the app registration.
  SCOPES: ['User.Read', 'Calendars.Read', 'Mail.Read', 'Tasks.Read', 'Chat.Read']
};
