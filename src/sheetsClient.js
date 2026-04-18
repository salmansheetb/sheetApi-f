/**
 * worker/src/sheetsClient.js
 * Google Sheets REST API client for Cloudflare Workers.
 * Uses Web Crypto API to sign JWTs — no googleapis dependency.
 */

const SHEETS_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';
const TOKEN_URL   = 'https://oauth2.googleapis.com/token';
const SCOPE       = 'https://www.googleapis.com/auth/spreadsheets';

// ── base64url helpers ─────────────────────────────────────────

function base64url(buffer) {
  const bytes = new Uint8Array(buffer);
  let bin = '';
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function base64urlFromString(str) {
  const bytes = new TextEncoder().encode(str);
  let bin = '';
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

// ── PEM → CryptoKey ───────────────────────────────────────────

async function importPrivateKey(pem) {
  const pemBody = pem
    .replace(/-----BEGIN PRIVATE KEY-----/, '')
    .replace(/-----END PRIVATE KEY-----/, '')
    .replace(/\s+/g, '');

  const binary = atob(pemBody);
  const bytes  = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

  return crypto.subtle.importKey(
    'pkcs8',
    bytes.buffer,
    { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
    false,
    ['sign']
  );
}

// ── JWT creation ──────────────────────────────────────────────

async function createJWT(clientEmail, privateKey) {
  const now = Math.floor(Date.now() / 1000);

  const header  = base64urlFromString(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payload = base64urlFromString(JSON.stringify({
    iss:   clientEmail,
    scope: SCOPE,
    aud:   TOKEN_URL,
    iat:   now,
    exp:   now + 3600,
  }));

  const signingInput = `${header}.${payload}`;
  const key          = await importPrivateKey(privateKey);
  const signature    = await crypto.subtle.sign(
    'RSASSA-PKCS1-v1_5',
    key,
    new TextEncoder().encode(signingInput)
  );

  return `${signingInput}.${base64url(signature)}`;
}

// ── OAuth2 token fetch ────────────────────────────────────────

async function fetchAccessToken(clientEmail, privateKey) {
  const jwt = await createJWT(clientEmail, privateKey);

  const res = await fetch(TOKEN_URL, {
    method:  'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body:    new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion:  jwt,
    }),
  });

  if (!res.ok) {
    const text = await res.text();
    const err  = new Error(`Failed to get Google access token: ${text}`);
    err.statusCode = 500;
    throw err;
  }

  const json = await res.json();
  // Return token + expiry timestamp (subtract 60s buffer)
  return {
    token:     json.access_token,
    expiresAt: Date.now() + (json.expires_in - 60) * 1000,
  };
}

// ── SheetsClient ──────────────────────────────────────────────

export class SheetsClient {
  constructor(clientEmail, privateKey) {
    this.clientEmail = clientEmail;
    // Normalize key: strip surrounding quotes, convert literal \n → real newlines
    this.privateKey  = privateKey
      .trim()
      .replace(/^"([\s\S]*)"$/, '$1')
      .replace(/\\n/g, '\n');
    this._tokenData = null;
  }

  async getToken() {
    // Refresh if missing or expired
    if (!this._tokenData || Date.now() >= this._tokenData.expiresAt) {
      this._tokenData = await fetchAccessToken(this.clientEmail, this.privateKey);
    }
    return this._tokenData.token;
  }

  async request(path, options = {}) {
    const token = await this.getToken();
    const url   = path.startsWith('http') ? path : `${SHEETS_BASE}${path}`;

    const res = await fetch(url, {
      ...options,
      headers: {
        Authorization:  `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...(options.headers ?? {}),
      },
    });

    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      const msg  = body?.error?.message ?? `Sheets API error ${res.status}`;
      const err  = new Error(msg);
      err.statusCode = res.status;
      throw err;
    }

    return res.json();
  }

  // ── Spreadsheet methods ───────────────────────────────────────

  spreadsheetGet(spreadsheetId) {
    return this.request(`/${spreadsheetId}`);
  }

  valuesGet(spreadsheetId, range) {
    const encoded = encodeURIComponent(range)
      .replace(/%27/g, "'")  // single-quote
      .replace(/%21/g, '!')  // exclamation
      .replace(/%3A/g, ':'); // colon (e.g. 1:1 row range)
    return this.request(`/${spreadsheetId}/values/${encoded}`);
  }

  valuesUpdate(spreadsheetId, range, body) {
    const encoded = encodeURIComponent(range)
      .replace(/%27/g, "'")
      .replace(/%21/g, '!')
      .replace(/%3A/g, ':');
    return this.request(
      `/${spreadsheetId}/values/${encoded}?valueInputOption=RAW`,
      { method: 'PUT', body: JSON.stringify(body) }
    );
  }

  valuesAppend(spreadsheetId, range, body) {
    const encoded = encodeURIComponent(range)
      .replace(/%27/g, "'")
      .replace(/%21/g, '!')
      .replace(/%3A/g, ':');
    return this.request(
      `/${spreadsheetId}/values/${encoded}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
      { method: 'POST', body: JSON.stringify(body) }
    );
  }

  valuesBatchUpdate(spreadsheetId, body) {
    return this.request(`/${spreadsheetId}/values:batchUpdate`, {
      method: 'POST',
      body:   JSON.stringify(body),
    });
  }

  batchUpdate(spreadsheetId, body) {
    return this.request(`/${spreadsheetId}:batchUpdate`, {
      method: 'POST',
      body:   JSON.stringify(body),
    });
  }
}

// ── Factory ───────────────────────────────────────────────────

/**
 * Returns a SheetsClient, reusing the same instance within a request
 * (env is stable per Worker request, so _sheetsClient lives for the request).
 */
export function getSheetsClient(env) {
  const { CLIENT_EMAIL, PRIVATE_KEY } = env;
  if (!CLIENT_EMAIL || !PRIVATE_KEY) {
    const err = new Error('Missing CLIENT_EMAIL or PRIVATE_KEY in Worker secrets.');
    err.statusCode = 500;
    throw err;
  }
  if (!env._sheetsClient) {
    env._sheetsClient = new SheetsClient(CLIENT_EMAIL, PRIVATE_KEY);
  }
  return env._sheetsClient;
}
