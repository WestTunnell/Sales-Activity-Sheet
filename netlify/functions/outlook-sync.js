// netlify/functions/outlook-sync.js
// Fetches recent emails + calendar events from Microsoft Graph
// Called from the SalesOS frontend

const { createClient } = require("@supabase/supabase-js");

async function refreshTokenIfNeeded(tokens, sb) {
  const CLIENT_ID     = process.env.OUTLOOK_CLIENT_ID;
  const CLIENT_SECRET = process.env.OUTLOOK_CLIENT_SECRET;
  const TENANT_ID     = process.env.OUTLOOK_TENANT_ID;

  const expiresAt = new Date(tokens.expires_at);
  const now = new Date();
  // Refresh if expiring within 5 minutes
  if ((expiresAt - now) > 5 * 60 * 1000) return tokens;

  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        refresh_token: tokens.refresh_token,
        grant_type:    "refresh_token",
        scope:         "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Calendars.Read offline_access"
      }).toString()
    }
  );

  const newTokens = await tokenRes.json();
  if (newTokens.error) throw new Error("Token refresh failed: " + newTokens.error_description);

  const updated = {
    access_token:  newTokens.access_token,
    refresh_token: newTokens.refresh_token || tokens.refresh_token,
    expires_at:    new Date(Date.now() + newTokens.expires_in * 1000).toISOString()
  };

  await sb.from("app_data").upsert({
    id: "outlook_tokens",
    payload: JSON.stringify(updated),
    updated_at: new Date().toISOString()
  });

  return updated;
}

exports.handler = async (event, context) => {
  const SUPABASE_URL = process.env.SUPABASE_URL;
  const SUPABASE_KEY = process.env.SUPABASE_SERVICE_KEY;

  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json"
  };

  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers, body: "" };
  }

  try {
    const sb = createClient(SUPABASE_URL, SUPABASE_KEY);

    // Load stored tokens
    const { data: tokenRow, error: tokenErr } = await sb
      .from("app_data")
      .select("payload")
      .eq("id", "outlook_tokens")
      .single();

    if (tokenErr || !tokenRow) {
      return { statusCode: 401, headers, body: JSON.stringify({ error: "not_connected" }) };
    }

    let tokens = JSON.parse(tokenRow.payload);
    tokens = await refreshTokenIfNeeded(tokens, sb);

    const graphHeaders = {
      Authorization: `Bearer ${tokens.access_token}`,
      "Content-Type": "application/json"
    };

    // Date range: last 90 days
    const since = new Date();
    since.setDate(since.getDate() - 90);
    const sinceISO = since.toISOString();

    // Fetch sent emails (from me) in last 90 days
    const emailRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/mailFolders/SentItems/messages`
      + `?$filter=sentDateTime ge ${sinceISO}`
      + `&$select=id,subject,sentDateTime,toRecipients,bodyPreview`
      + `&$top=50&$orderby=sentDateTime desc`,
      { headers: graphHeaders }
    );
    const emailData = await emailRes.json();
    const emails = (emailData.value || []).map(m => ({
      id: m.id,
      subject: m.subject,
      date: m.sentDateTime ? m.sentDateTime.slice(0, 10) : null,
      to: (m.toRecipients || []).map(r => ({
        name: r.emailAddress.name,
        email: r.emailAddress.address
      })),
      preview: m.bodyPreview ? m.bodyPreview.slice(0, 200) : "",
      source: "email_sent"
    }));

    // Fetch received emails in last 90 days
    const inboxRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages`
      + `?$filter=receivedDateTime ge ${sinceISO}`
      + `&$select=id,subject,receivedDateTime,from,bodyPreview`
      + `&$top=50&$orderby=receivedDateTime desc`,
      { headers: graphHeaders }
    );
    const inboxData = await inboxRes.json();
    const received = (inboxData.value || []).map(m => ({
      id: m.id,
      subject: m.subject,
      date: m.receivedDateTime ? m.receivedDateTime.slice(0, 10) : null,
      from: m.from ? {
        name: m.from.emailAddress.name,
        email: m.from.emailAddress.address
      } : null,
      preview: m.bodyPreview ? m.bodyPreview.slice(0, 200) : "",
      source: "email_received"
    }));

    // Fetch calendar events in last 90 days
    const calRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/calendarView`
      + `?startDateTime=${sinceISO}`
      + `&endDateTime=${new Date().toISOString()}`
      + `&$select=id,subject,start,end,attendees,bodyPreview,location`
      + `&$top=50&$orderby=start/dateTime desc`,
      { headers: graphHeaders }
    );
    const calData = await calRes.json();
    const events = (calData.value || []).map(e => ({
      id: e.id,
      subject: e.subject,
      date: e.start && e.start.dateTime ? e.start.dateTime.slice(0, 10) : null,
      location: e.location && e.location.displayName ? e.location.displayName : "",
      attendees: (e.attendees || []).map(a => ({
        name: a.emailAddress.name,
        email: a.emailAddress.address
      })),
      preview: e.bodyPreview ? e.bodyPreview.slice(0, 200) : "",
      source: "calendar"
    }));

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ emails, received, events, synced_at: new Date().toISOString() })
    };

  } catch (err) {
    console.error("Sync error:", err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: err.message })
    };
  }
};
