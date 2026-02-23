const { getItem, setItem, deleteItem, randomId } = require("./_store");

exports.handler = async (event) => {
  const code = event.queryStringParameters?.code;
  const state = event.queryStringParameters?.state;
  const error = event.queryStringParameters?.error;

  if (error) {
    return {
      statusCode: 400,
      body: JSON.stringify({ ok: false, error }),
    };
  }

  if (!code || !state) {
    return {
      statusCode: 400,
      body: JSON.stringify({ ok: false, error: "Missing code or state." }),
    };
  }

  const stateEntry = await getItem(`state:${state}`);
  if (!stateEntry) {
    return {
      statusCode: 400,
      body: JSON.stringify({ ok: false, error: "Invalid or expired state." }),
    };
  }
  await deleteItem(`state:${state}`);

  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const redirectUri = process.env.GOOGLE_REDIRECT_URI;

  if (!clientId || !clientSecret || !redirectUri) {
    return {
      statusCode: 500,
      body: JSON.stringify({ ok: false, error: "Missing Google OAuth environment variables." }),
    };
  }

  const tokenResponse = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      code,
      client_id: clientId,
      client_secret: clientSecret,
      redirect_uri: redirectUri,
      grant_type: "authorization_code",
    }),
  });

  const tokenBody = await tokenResponse.json();
  if (!tokenResponse.ok) {
    return {
      statusCode: 400,
      body: JSON.stringify({ ok: false, error: tokenBody }),
    };
  }

  const sessionId = randomId();
  await setItem(`session:${sessionId}`, tokenBody, 3600);

  const proto = event.headers["x-forwarded-proto"] || "https";
  const host = event.headers.host;
  const redirectUrl = `${proto}://${host}/#/app?oauth=success`;

  return {
    statusCode: 302,
    headers: {
      Location: redirectUrl,
      "Set-Cookie": `saai_session=${sessionId}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=3600`,
      "Cache-Control": "no-store",
    },
  };
};
