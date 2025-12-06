
// ===== Replace with your values =====
const tenantId   = "4add8a0e-95e4-4321-95e1-352a5ce5cfd8";          // e.g., "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
const clientId   = "07a53270-096a-4b58-b3ac-b62bf0de8665";          // Application (client) ID of your SPA app
const authority  = `https://login.microsoftonline.com/${tenantId}`;
// If you support MSA or multi-tenant, adjust authority accordingly

// GitHub Pages URL you set as Redirect URI (must exactly match)
const redirectUri = "https://github.com/gentarom/GSAWebCategoryCheck";

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
