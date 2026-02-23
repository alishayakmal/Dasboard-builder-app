const { getItem } = require("./_store");

function parseCookies(cookieHeader) {
  const cookies = {};
  if (!cookieHeader) return cookies;
  cookieHeader.split(";").forEach((pair) => {
    const [key, ...rest] = pair.trim().split("=");
    cookies[key] = decodeURIComponent(rest.join("="));
  });
  return cookies;
}

exports.handler = async (event) => {
  const cookies = parseCookies(event.headers.cookie || "");
  const sessionId = cookies.saai_session;
  if (!sessionId) {
    return {
      statusCode: 401,
      body: JSON.stringify({ ok: false, error: "Not authenticated." }),
    };
  }

  const tokens = await getItem(`session:${sessionId}`);
  if (!tokens?.access_token) {
    return {
      statusCode: 401,
      body: JSON.stringify({ ok: false, error: "Session expired." }),
    };
  }

  const spreadsheetId = event.queryStringParameters?.spreadsheetId;
  const range = event.queryStringParameters?.range;
  if (!spreadsheetId || !range) {
    return {
      statusCode: 400,
      body: JSON.stringify({ ok: false, error: "Missing spreadsheetId or range." }),
    };
  }

  const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`;
  const response = await fetch(apiUrl, {
    headers: { Authorization: `Bearer ${tokens.access_token}` },
  });

  const body = await response.text();
  if (!response.ok) {
    return {
      statusCode: response.status,
      body,
    };
  }

  return {
    statusCode: 200,
    headers: { "Content-Type": "application/json" },
    body,
  };
};
