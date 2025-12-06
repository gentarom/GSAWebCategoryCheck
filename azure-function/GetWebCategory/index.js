
// index.js
import fetch from "node-fetch";
import { ConfidentialClientApplication } from "@azure/msal-node";

// ---- Required environment variables (fail fast if any missing) ----
const requiredEnv = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "ALLOWED_ORIGIN"];
for (const k of requiredEnv) {
  if (!process.env[k]) throw new Error(`Missing environment variable: ${k}`);
}

const tenantId      = process.env.TENANT_ID;      // e.g., "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
const clientId      = process.env.CLIENT_ID;      // App registration (client) ID
const clientSecret  = process.env.CLIENT_SECRET;  // Secure Value of the client secret
const allowedOrigin = process.env.ALLOWED_ORIGIN; // e.g., "https://gentarom.github.io/GSAWebCategoryCheck/"

// ---- MSAL confidential client (client credentials) ----
const cca = new ConfidentialClientApplication({
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret
  }
});

// ---- CORS helper ----
function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": allowedOrigin,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };
}

// ---- Azure Function entry point ----
export default async function (context, req) {
  // Handle browser preflight
  if (req.method === "OPTIONS") {
    context.res = { status: 204, headers: corsHeaders() };
    return;
  }

  try {
    // 1) Validate input
    const inputUrl = (req.body && req.body.url ? String(req.body.url) : "").trim();
    if (!inputUrl) {
      context.res = { status: 400, headers: corsHeaders(), body: { error: "Missing 'url      context.res = { status: 400, headers: corsHeaders(), body: { error: "Missing 'url'." } };
      return;
    }

    // 2) Normalize for the function parameter:
    //    - Remove scheme (http/https)
    //    - Remove leading slashes
    //    - Escape single quotes for OData string literal safety
    const withoutScheme = inputUrl.replace(/^https?:\/\//i, "");
    const normalized    = withoutScheme.replace(/^\/+/, "");
    const safeLiteral   = normalized.replace(/'/g, "''"); // OData string escape

    // 3) Acquire app-only token
    let accessToken;
    try {
      const tokenResponse = await cca.acquireTokenByClientCredential({
        scopes: ["https://graph.microsoft.com/.default"]
      });
      accessToken = tokenResponse.accessToken;
    } catch (e) {
      context.log.error("Token acquisition failed:", e);
      context.res = {
        status: 500,
        headers: corsHeaders(),
        body: { error: "Token acquisition failed", detail: String(e) }
      };
      return;
    }

    // 4) Build the endpoint EXACTLY as your tenant accepts:
    //    GET /beta/networkAccess/connectivity/microsoft.graph.networkaccess.getWebCategoryByUrl(url='host[/path]')
    const endpoint =
      "https://graph.microsoft.com/beta/networkAccess/connectivity/" +
      `microsoft.graph.networkaccess.getWebCategoryByUrl(url='${safeLiteral}')`;

    context.log(`Calling Graph (GET-literal-no-scheme): ${endpoint}`);

    // 5) Call Graph
    const graphRes = await fetch(endpoint, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    const raw = await graphRes.text();
    context.log(`Graph status: ${graphRes.status}; body length: ${raw ? raw.length : 0}`);

    // 6) Parse if JSON; otherwise, return raw payload for visibility
    let body;
    try {
      body = raw ? JSON.parse(raw) : {};
    } catch {
      body = { raw, note: "Response was not valid JSON." };
    }

    context.res = { status: graphRes.status, headers: corsHeaders(), body };
  } catch (err) {
    context.log.error("Unhandled server error:", err);
    context.res = {
      status: 500,
      headers: corsHeaders(),
      body: { error: "Unhandled server error", detail: String(err) }
    };
  }
}
