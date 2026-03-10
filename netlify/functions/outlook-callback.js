// netlify/functions/outlook-callback.js
// Exchanges the auth code for tokens and stores them in Supabase

const { createClient } = require("@supabase/supabase-js");

exports.handler = async (event, context) => {
  const CLIENT_ID     = process.env.OUTLOOK_CLIENT_ID;
  const CLIENT_SECRET = process.env.OUTLOOK_CLIENT_SECRET;
  const TENANT_ID     = process.env.OUTLOOK_TENANT_ID;
  const REDIRECT_URI  = process.env.OUTLOOK_REDIRECT_URI || "https://faircosalesactivity.netlify.app/.netlify/functions/outlook-callback";
  const SUPABASE_URL  = process.env.SUPABASE_URL;
  const SUPABASE_KEY  = process.env.SUPABASE_SERVICE_KEY; // use service key server-side

  const code = event.queryStringParameters && event.queryStringParameters.code;
  if (!code) {
    return { statusCode: 400, body: "Missing auth code" };
  }

  try {
    // Exchange code for tokens
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id:     CLIENT_ID,
          client_secret: CLIENT_SECRET,
          code:          code,
          redirect_uri:  REDIRECT_URI,
          grant_type:    "authorization_code"
        }).toString()
      }
    );

    const tokens = await tokenRes.json();
    if (tokens.error) throw new Error(tokens.error_description || tokens.error);

    // Store tokens in Supabase
    const sb = createClient(SUPABASE_URL, SUPABASE_KEY);
    const expiresAt = new Date(Date.now() + tokens.expires_in * 1000).toISOString();

    await sb.from("app_data").upsert({
      id: "outlook_tokens",
      payload: JSON.stringify({
        access_token:  tokens.access_token,
        refresh_token: tokens.refresh_token,
        expires_at:    expiresAt
      }),
      updated_at: new Date().toISOString()
    });

    // Redirect back to app with success flag
    return {
      statusCode: 302,
      headers: { Location: "https://faircosalesactivity.netlify.app?outlook=connected" },
      body: ""
    };
  } catch (err) {
    console.error("Outlook callback error:", err);
    return {
      statusCode: 302,
      headers: { Location: "https://faircosalesactivity.netlify.app?outlook=error" },
      body: ""
    };
  }
};
