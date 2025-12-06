import fetch from "node-fetch";
import { ConfidentialClientApplication } from "@azure/msal-node";

// Read secrets from environment (configure in Azure → Function App → Configuration)
const tenantId      = process.env.TENANT_ID;      // e.g., aaaa-bbbb-cccc-dddd
const clientId      = process.env.CLIENT_ID;
const clientSecret  = process.env.CLIENT_SECRET;  // or use certificates
const allowedOrigin = process.env.ALLOWED_ORIGIN || "*"; // e.g., https://gentarom.github.io

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

    // Acquire Graph app-only token
    const tokenResponse = await cca.acquireTokenByClientCredential({
      scopes: [ "https://graph.microsoft.com/.default" ]
    });
    const accessToken = tokenResponse.accessToken;

    // 1) Try function-style GET (your requested endpoint)
    const encoded = encodeURIComponent(url);
    const getEndpoint =
      `https://graph.microsoft.com/beta/networkaccess/connectivity/` +
      `microsoft.graph.networkaccess.getWebCategoryByUrl(url='${encoded}')`;

    let graphRes = await fetch(getEndpoint, {
      method: "GET",
      headers: { "Authorization": `Bearer ${accessToken}`, "Accept": "application/json" }
    });

    // 2) If GET fails (404/400/403, etc.), fall back to documented POST action
    if (!graphRes.ok) {
      const postEndpoint = "https://graph.microsoft.com/beta/networkAccess/categorizeWebUrl";
      graphRes = await fetch(postEndpoint, {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Content-Type": "application/json",
          "Accept": "application/json"
        },
               body: JSON.stringify({ url })
      });
    }

    const text = await graphRes.text();
    const body = text ? JSON.parse(text) : {};
    context.res = { status: graphRes.status, headers: corsHeaders(), body };
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, headers: corsHeaders(), body: { error: "Server error", detail: `${err}` } };
  }
}
