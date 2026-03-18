const express = require('express');
const fetch = (...args) => import('node-fetch').then(m => m.default(...args));
const app = express();
app.use(express.json());

const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const REFRESH_TOKEN = process.env.AZURE_REFRESH_TOKEN;
const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';
const SCOPE = 'openid profile offline_access User.Read Mail.ReadWrite Mail.Send';

async function refreshAccessToken() {
const params = new URLSearchParams();
params.append('client_id', CLIENT_ID);
params.append('grant_type', 'refresh_token');
params.append('refresh_token', REFRESH_TOKEN);
params.append('scope', SCOPE);

const r = await fetch(TOKEN_ENDPOINT, { method: 'POST', body: params });
if (!r.ok) {
const t = await r.text();
throw new Error('Token refresh failed: ' + t);
}
return r.json();
}

app.get('/messages', async (req, res) => {
try {
const tokens = await refreshAccessToken();
const access = tokens.access_token;
const q = 'https://graph.microsoft.com/v1.0/me/messages?$top=20&$select=receivedDateTime,from,subject';
const r = await fetch(q, { headers: { Authorization: Bearer ${access} } });
const j = await r.json();
res.json(j);
} catch (err) {
console.error(err);
res.status(500).json({ error: String(err) });
}
});

app.get('/message/:id', async (req, res) => {
try {
const tokens = await refreshAccessToken();
const access = tokens.access_token;
const q = https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(req.params.id)};
const r = await fetch(q, { headers: { Authorization: Bearer ${access} } });
const j = await r.json();
res.json(j);
} catch (err) {
console.error(err);
res.status(500).json({ error: String(err) });
}
});

app.post('/send', async (req, res) => {
try {
const { to, subject, body } = req.body;
if (!to || !subject || !body) return res.status(400).json({ error: 'to, subject, body required' });

const tokens = await refreshAccessToken();
const access = tokens.access_token;
const payload = {
  message: {
    subject,
    body: { contentType: 'Text', content: body },
    toRecipients: [{ emailAddress: { address: to } }]
  }
};

const r = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
  method: 'POST',
  headers: { Authorization: `Bearer ${access}`, 'Content-Type': 'application/json' },
  body: JSON.stringify(payload)
});

if (!r.ok) {
  const t = await r.text();
  return res.status(500).json({ error: t });
}
res.json({ status: 'sent' });
} catch (err) {
console.error(err);
res.status(500).json({ error: String(err) });
}
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log('Worker listening on', port));
