
import fetch from "node-fetch";
import { ConfidentialClientApplication } from "@azure/msal-node";

// === Env vars (fail fast if missing) ===
const requiredEnv = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "ALLOWED_ORIGIN"];
for (const k of requiredEnv) {
  if (!process.env[k]) throw new Error(`Missing environment variable: ${k}`);
}
const tenantId      = process.env.TENANT_ID;
const clientId      = process.env.CLIENT_ID;
const clientSecret  = process.env.CLIENT_SECRET;
const allowedOrigin = process.env.ALLOWED_ORIGIN;

// === MSAL confidential client ===
const cca = new ConfidentialClientApplication({
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret
  }
});

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": allowedOrigin,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };
}

export default async function (context, req) {
  // CORS preflight
  if (req.method === "OPTIONS") {
    context.res = { status: 204, headers: corsHeaders() };
    return;
  }

  try {
    const url = (req.body && req.body.url || "").trim();
    if (!url) {
      context.res = { status: 400, headers: corsHeaders(), body: { error: "Missing 'url'." } };
      return;
    }

    // Acquire app-only token
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

    // Build function-style GET with correct alias syntax (value quoted after ?@x=)
    const endpoint =
      "https://graph.microsoft.com/beta/networkaccess/connectivity/" +
      "microsoft.graph.networkaccess.getWebCategoryByUrl(url=@x)?@x='" +
      encodeURIComponent(url) + "'";

    context.log(`Calling Graph: ${endpoint}`);

    const graphRes = await fetch(endpoint, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    const raw = await graphRes.text();
    context.log(`Graph status: ${graphRes.status}; body length: ${raw?.length || 0}`);

    // Parse if JSON; otherwise return raw payload for visibility
    let body;
    try {
      body = raw ? JSON.parse(raw) : {};
    } catch {
      body = { raw, note: "Response was not valid JSON." };
    }

    context.res = { status: graphRes.status, headers: corsHeaders(), body };
  } catch (err) {
    context.log.error("Unhandled server error:", err);
       context.res = { status: 500, headers: corsHeaders(), body: { error: "Unhandled server error", detail: String(err) } };
  }
}
