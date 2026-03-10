// netlify/functions/outlook-auth.js
// Redirects the user to Microsoft login

exports.handler = async (event, context) => {
  const CLIENT_ID   = process.env.OUTLOOK_CLIENT_ID;
  const TENANT_ID   = process.env.OUTLOOK_TENANT_ID;
  const REDIRECT_URI = process.env.OUTLOOK_REDIRECT_URI || "https://faircosalesactivity.netlify.app/.netlify/functions/outlook-callback";

  const scopes = [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Calendars.Read",
    "offline_access"
  ].join(" ");

  const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`
    + `?client_id=${CLIENT_ID}`
    + `&response_type=code`
    + `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}`
    + `&scope=${encodeURIComponent(scopes)}`
    + `&response_mode=query`;

  return {
    statusCode: 302,
    headers: { Location: authUrl },
    body: ""
  };
};
