
// ===== Replace with your values =====
const tenantId   = "<your-tenant-id>";          // e.g., "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
const clientId   = "<your-client-id>";          // Application (client) ID of your SPA app
const authority  = `https://login.microsoftonline.com/${tenantId}`;
// If you support MSA or multi-tenant, adjust authority accordingly

// GitHub Pages URL you set as Redirect URI (must exactly match)
const redirectUri = "https://<your-github-username>.github.io/<repo-name>/";

// Graph scopes required for the API call (delegated)
const graphScopes = ["NetworkAccess.Read.All"]; // admin consent required

// MSAL instance configuration
const msalConfig = {
  auth: {
    clientId,
    authority,
    redirectUri
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
};

window.__msalConfig__ = { msalConfig, graphScopes };
