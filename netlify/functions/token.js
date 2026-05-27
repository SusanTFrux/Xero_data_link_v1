/*
  netlify/functions/token.js
  Exchanges the OAuth auth code for an access token.
  Runs server-side on Netlify to bypass CORS restrictions on Xero's token endpoint.
  Client credentials stay in Netlify environment variables — never sent to browser.
*/

exports.handler = async function(event) {
  // CORS headers
  const headers = {
    "Access-Control-Allow-Origin":  "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Content-Type": "application/json"
  };

  // Handle CORS preflight
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers, body: "" };
  }

  if (event.httpMethod !== "POST") {
    return { statusCode: 405, headers, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  try {
    const body = JSON.parse(event.body || "{}");
    const { code, verifier, redirect_uri } = body;

    if (!code || !verifier || !redirect_uri) {
      return {
        statusCode: 400, headers,
        body: JSON.stringify({ error: "Missing required fields: code, verifier, redirect_uri" })
      };
    }

    const clientId     = process.env.XERO_CLIENT_ID;
    const clientSecret = process.env.XERO_CLIENT_SECRET;

    if (!clientId || !clientSecret) {
      return {
        statusCode: 500, headers,
        body: JSON.stringify({ error: "Server missing XERO_CLIENT_ID or XERO_CLIENT_SECRET environment variables." })
      };
    }

    const basicAuth = Buffer.from(clientId + ":" + clientSecret).toString("base64");

    const formBody = new URLSearchParams({
      grant_type:    "authorization_code",
      code:          code,
      redirect_uri:  redirect_uri,
      code_verifier: verifier
    }).toString();

    const xeroResp = await fetch("https://identity.xero.com/connect/token", {
      method: "POST",
      headers: {
        "Content-Type":  "application/x-www-form-urlencoded",
        "Authorization": "Basic " + basicAuth
      },
      body: formBody
    });

    const xeroData = await xeroResp.json();

    if (!xeroResp.ok) {
      return {
        statusCode: xeroResp.status, headers,
        body: JSON.stringify({ error: xeroData.error || "Token exchange failed", detail: xeroData })
      };
    }

    return { statusCode: 200, headers, body: JSON.stringify(xeroData) };

  } catch(err) {
    return {
      statusCode: 500, headers,
      body: JSON.stringify({ error: err.message })
    };
  }
};
