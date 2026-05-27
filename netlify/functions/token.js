/*
  netlify/functions/token.js
  Exchanges the OAuth auth code for a Xero access token.
  Uses Node's built-in https module — no external dependencies needed.
*/

const https = require("https");

// Helper: make an HTTPS POST request and return parsed JSON
function httpsPost(hostname, path, headers, body) {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: hostname,
      path: path,
      method: "POST",
      headers: {
        ...headers,
        "Content-Length": Buffer.byteLength(body)
      }
    };

    const req = https.request(options, (res) => {
      let data = "";
      res.on("data", chunk => data += chunk);
      res.on("end", () => {
        try {
          resolve({ status: res.statusCode, body: JSON.parse(data) });
        } catch(e) {
          resolve({ status: res.statusCode, body: data });
        }
      });
    });

    req.on("error", reject);
    req.write(body);
    req.end();
  });
}

exports.handler = async function(event) {

  // CORS headers — allow requests from any origin (required for Excel add-in)
  const cors = {
    "Access-Control-Allow-Origin":  "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Content-Type": "application/json"
  };

  // Handle CORS preflight request
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: cors, body: "" };
  }

  // Only allow POST
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, headers: cors, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  try {
    // Parse request body
    const body = JSON.parse(event.body || "{}");
    const { code, verifier, redirect_uri } = body;

    if (!code || !verifier || !redirect_uri) {
      return {
        statusCode: 400, headers: cors,
        body: JSON.stringify({ error: "Missing: code, verifier, or redirect_uri" })
      };
    }

    // Get credentials from Netlify environment variables
    const clientId     = process.env.XERO_CLIENT_ID;
    const clientSecret = process.env.XERO_CLIENT_SECRET;

    if (!clientId || !clientSecret) {
      return {
        statusCode: 500, headers: cors,
        body: JSON.stringify({ error: "Missing XERO_CLIENT_ID or XERO_CLIENT_SECRET in Netlify environment variables" })
      };
    }

    // Build Basic Auth header
    const basicAuth = Buffer.from(clientId + ":" + clientSecret).toString("base64");

    // Build form body for Xero token request
    const formBody = new URLSearchParams({
      grant_type:    "authorization_code",
      code:          code,
      redirect_uri:  redirect_uri,
      code_verifier: verifier
    }).toString();

    // Call Xero token endpoint using https module
    const result = await httpsPost(
      "identity.xero.com",
      "/connect/token",
      {
        "Content-Type":  "application/x-www-form-urlencoded",
        "Authorization": "Basic " + basicAuth
      },
      formBody
    );

    if (result.status !== 200) {
      return {
        statusCode: result.status, headers: cors,
        body: JSON.stringify({ error: "Xero token exchange failed", detail: result.body })
      };
    }

    return {
      statusCode: 200, headers: cors,
      body: JSON.stringify(result.body)
    };

  } catch(err) {
    return {
      statusCode: 500, headers: cors,
      body: JSON.stringify({ error: err.message })
    };
  }
};
