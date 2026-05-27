/*
  netlify/functions/xero.js
  General-purpose Xero API proxy.
  Routes all Xero API calls through Netlify to avoid CORS restrictions
  in Excel's sandboxed browser context.

  Usage: POST with { endpoint, token, tenantId, params }
  Returns the Xero API response as JSON.
*/

const https = require("https");

function httpsGet(hostname, path, headers) {
  return new Promise((resolve, reject) => {
    const options = { hostname, path, method: "GET", headers };
    const req = https.request(options, (res) => {
      let data = "";
      res.on("data", chunk => data += chunk);
      res.on("end", () => {
        try { resolve({ status: res.statusCode, body: JSON.parse(data) }); }
        catch(e) { resolve({ status: res.statusCode, body: data }); }
      });
    });
    req.on("error", reject);
    req.end();
  });
}

exports.handler = async function(event) {
  const cors = {
    "Access-Control-Allow-Origin":  "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Content-Type": "application/json"
  };

  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers: cors, body: "" };
  if (event.httpMethod !== "POST") return { statusCode: 405, headers: cors, body: JSON.stringify({ error: "Method not allowed" }) };

  try {
    const body = JSON.parse(event.body || "{}");
    const { endpoint, token, tenantId, params } = body;

    if (!endpoint || !token) {
      return { statusCode: 400, headers: cors, body: JSON.stringify({ error: "Missing endpoint or token" }) };
    }

    // Build path — endpoint can be "connections" or "api/Reports/ProfitAndLoss" etc.
    let hostname, path;

    if (endpoint === "connections") {
      // Tenant list endpoint
      hostname = "api.xero.com";
      path = "/connections";
    } else {
      // Standard API endpoint
      hostname = "api.xero.com";
      path = "/api.xro/2.0/" + endpoint;
      if (params) path += "?" + new URLSearchParams(params).toString();
    }

    const reqHeaders = {
      "Authorization": "Bearer " + token,
      "Accept": "application/json"
    };

    if (tenantId) reqHeaders["Xero-Tenant-Id"] = tenantId;

    const result = await httpsGet(hostname, path, reqHeaders);

    return { statusCode: result.status, headers: cors, body: JSON.stringify(result.body) };

  } catch(err) {
    return { statusCode: 500, headers: cors, body: JSON.stringify({ error: err.message }) };
  }
};
