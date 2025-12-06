
import fetch from "node-fetch";
import { ConfidentialClientApplication } from "@azure/msal-node";

// ---- Env vars (fail fast) ----
const requiredEnv = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "ALLOWED_ORIGIN"];
for (const k of requiredEnv) {
  if (!process.env[k]) throw new Error(`Missing environment variable: ${k}`);
}
const tenantId      = process.env.TENANT_ID;
const clientId      = process.env.CLIENT_ID;
const clientSecret  = process.env.CLIENT_SECRET;
const allowedOrigin = process.env.ALLOWED_ORIGIN;

// ---- MSAL confidential client ----
const cca = new ConfidentialClientApplication({
  auth: { clientId, authority: `https://login.microsoftonline.com/${tenantId}`, clientSecret }
});

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": allowedOrigin,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };
}

// Helper: call Graph and return {status, raw, json?, endpoint}
async function callGraphGET(endpoint, accessToken, context) {
  context.log(`Calling Graph (GET): ${endpoint}`);
  const res = await fetch(endpoint, {
    method: "GET",
    headers: { "Authorization": `Bearer ${accessToken}`, "Accept": "application/json" }
  });
  const raw = await res.text();
  let json; try { json = raw ? JSON.parse(raw) : undefined; } catch { /* keep raw */ }
  context.log(`Graph GET status: ${res.status}; body length: ${raw?.length || 0}`);
  return { status: res.status, raw, json, endpoint };
}

async function callGraphPOST(url, accessToken, context) {
  const endpoint = "https://graph.microsoft.com/beta/networkAccess/categorizeWebUrl";
  context.log(`Calling Graph (POST): ${endpoint} body: ${url}`);
  const res = await fetch(endpoint, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    body: JSON.stringify({ url })
  });
  const raw = await res.text();
  let json; try { json = raw ? JSON.parse(raw) : undefined; } catch { /* keep raw */ }
  context.log(`Graph POST status: ${res.status}; body length: ${raw?.length || 0}`);
  return { status: res.status, raw, json, endpoint };
}

export default async function (context, req) {
  // Preflight
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
      const tokenResponse = await cca.acquireTokenByClientCredential({ scopes: ["https://graph.microsoft.com/.default"] });
      accessToken = tokenResponse.accessToken;
    } catch (e) {
      context.log.error("Token acquisition failed:", e);
      context.res = { status: 500, headers: corsHeaders(), body: { error: "Token acquisition failed", detail: String(e) } };
      return;
    }

    // Try GET variants in order:
    const attempts = [
      // A1: alias (camel-case path, UNencoded, quoted value)
      `https://graph.microsoft.com/beta/networkAccess/connectivity/microsoft.graph.networkaccess.getWebCategoryByUrl(url=@x)?@x='${url}'`,
      // A2: alias (lower-case path, UNencoded, quoted value)
      `https://graph.microsoft.com/beta/networkaccess/connectivity/microsoft.graph.networkaccess.getWebCategoryByUrl(url=@x)?@x='${url}'`,
      // B1: literal (camel-case path, UNencoded)
      `https://graph.microsoft.com/beta/networkAccess/connectivity/microsoft.graph.networkaccess.getWebCategoryByUrl(url='${url}')`,
      // B2: literal (lower-case path, UNencoded)
      `https://graph.microsoft.com/beta/networkaccess/connectivity/microsoft.graph.networkaccess.getWebCategoryByUrl(url='${url}')`
    ];

    let last = null;
    for (const ep of attempts) {
      const r = await callGraphGET(ep, accessToken, context);
      last = r;
      if (r.status >= 200 && r.status < 300) {
        context.res = { status: r.status, headers: corsHeaders(), body: r.json ?? { raw: r.raw } };
        return;
      }
      // Some tenants misreport with 500 but also accept encoded; try encoded once per branch
      if (r.status === 400 || r.status === 500) {
        // encoded value variant of current ep (keep quotes around value)
        const encodedVal = encodeURIComponent(url);
        const encodedEp = ep.replace(`'${url}'`, `'${encodedVal}'`);
        const re = await callGraphGET(encodedEp, accessToken, context);
        if (re.status >= 200 && re.status < 300) {
          context.res = { status: re.status, headers: corsHeaders(), body: re.json ?? { raw: re.raw } };
          return;
        }
        last = re;
      }
    }

    // Final fallback: POST action (if available in your tenant)
    const post = await callGraphPOST(url, accessToken, context);
    if (post.status >= 200 && post.status < 300) {
      context.res = { status: post.status, headers: corsHeaders(), body: post.json ?? { raw: post.raw } };
      return;
    }

    // None succeeded — return the last Graph payload for visibility
    context.res = {
      status: last?.status ?? post.status,
      headers: corsHeaders(),
      body: last?.json ?? post.json ?? { raw: (last?.raw ?? post.raw), note: "Graph call did not succeed." }
    };
  } catch (err) {
    context.log.error("Unhandled server error:", err);
    context.res = { status: 500, headers: corsHeaders(), body: { error: "Unhandled server error", detail: String(err) } };
  }
}
