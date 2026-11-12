// Delegated Microsoft Entra sign-in and Graph token storage. Graph tokens are kept per Entra user (oid)
// in KV, encrypted with TOKEN_ENC_KEY, and refreshed (with the rotated refresh token stored) on demand.

const authority = (env) => `https://login.microsoftonline.com/${env.MS365_TENANT_ID}/oauth2/v2.0`;
const scopes = (env) => env.GRAPH_SCOPES || 'https://graph.microsoft.com/.default offline_access openid profile';
export const redirectUri = (req) => new URL('/auth/callback', req.url).toString();

const b64url = (bytes) => btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const fromB64url = (s) => Uint8Array.from(atob(s.replace(/-/g, '+').replace(/_/g, '/')), (c) => c.charCodeAt(0));

export async function pkcePair() {
  const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  return { verifier, challenge };
}

export function authorizeUrl(env, req, { state, codeChallenge }) {
  const url = new URL(`${authority(env)}/authorize`);
  url.search = new URLSearchParams({
    client_id: env.MS365_APP_CLIENT_ID,
    response_type: 'code',
    response_mode: 'query',
    redirect_uri: redirectUri(req),
    scope: scopes(env),
    state,
    code_challenge: codeChallenge,
    code_challenge_method: 'S256',
    prompt: 'select_account',
  }).toString();
  return url.toString();
}

export class EntraError extends Error {
  constructor(body) {
    super(`Microsoft sign-in error: ${body.error}: ${body.error_description?.split('\r\n')[0]}`);
    this.code = body.error;
  }
}

async function tokenRequest(env, params) {
  const res = await fetch(`${authority(env)}/token`, {
    method: 'POST',
    body: new URLSearchParams({ client_id: env.MS365_APP_CLIENT_ID, client_secret: env.MS365_APP_CLIENT_SECRET, scope: scopes(env), ...params }),
  });
  const body = await res.json();
  if (!res.ok) throw new EntraError(body);
  return body;
}

export const redeemCode = (env, req, code, codeVerifier) =>
  tokenRequest(env, { grant_type: 'authorization_code', code, redirect_uri: redirectUri(req), code_verifier: codeVerifier });

export function idTokenClaims(idToken) {
  return JSON.parse(new TextDecoder().decode(fromB64url(idToken.split('.')[1])));
}

let keyPromise;
const encKey = (env) =>
  (keyPromise ??= crypto.subtle.importKey('raw', fromB64url(env.TOKEN_ENC_KEY), 'AES-GCM', false, ['encrypt', 'decrypt']));

async function seal(env, value) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const data = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await encKey(env), new TextEncoder().encode(JSON.stringify(value)));
  return `${b64url(iv)}.${b64url(data)}`;
}

async function unseal(env, sealed) {
  const [iv, data] = sealed.split('.').map(fromB64url);
  return JSON.parse(new TextDecoder().decode(await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, await encKey(env), data)));
}

const userKey = (userId) => `graph-user:${userId}`;

async function saveUser(env, userId, user) {
  await env.OAUTH_KV.put(userKey(userId), JSON.stringify({ upn: user.upn, signedInAt: user.signedInAt, refreshedAt: user.refreshedAt, sealed: await seal(env, user) }));
}

async function loadUser(env, userId) {
  const raw = await env.OAUTH_KV.get(userKey(userId));
  return raw ? unseal(env, JSON.parse(raw).sealed) : null;
}

// Stores the account's Graph tokens only after allow(account) passes, so refused accounts leave nothing behind.
export async function storeUserTokens(env, tokens, allow) {
  const claims = idTokenClaims(tokens.id_token);
  const account = { userId: claims.oid, upn: (claims.preferred_username || claims.upn || '').toLowerCase(), tenantId: claims.tid };
  const refusal = allow(account);
  if (refusal) return { refusal };
  const user = {
    upn: (claims.preferred_username || claims.upn || '').toLowerCase(),
    name: claims.name,
    tenantId: claims.tid,
    refreshToken: tokens.refresh_token,
    accessToken: tokens.access_token,
    accessExpiresAt: Date.now() + (tokens.expires_in - 120) * 1000,
    signedInAt: new Date().toISOString(),
  };
  await saveUser(env, claims.oid, user);
  return account;
}

const refreshing = new Map();

// A valid Graph access token for the user, refreshing it when it has expired.
export async function graphAccessToken(env, userId) {
  const user = await loadUser(env, userId);
  if (!user) throw new Error('No Microsoft account is linked to this connection; reconnect the MCP server.');
  if (user.accessToken && user.accessExpiresAt > Date.now()) return user.accessToken;
  if (!refreshing.has(userId)) {
    refreshing.set(
      userId,
      tokenRequest(env, { grant_type: 'refresh_token', refresh_token: user.refreshToken })
        .then(async (tokens) => {
          user.accessToken = tokens.access_token;
          user.accessExpiresAt = Date.now() + (tokens.expires_in - 120) * 1000;
          if (tokens.refresh_token) user.refreshToken = tokens.refresh_token;
          user.refreshedAt = new Date().toISOString();
          await saveUser(env, userId, user);
          return user.accessToken;
        })
        .catch((err) => {
          if (err.code === 'invalid_grant') throw new Error('Microsoft sign-in has expired or was revoked; reconnect the MCP server.');
          throw err;
        })
        .finally(() => refreshing.delete(userId)),
    );
  }
  return refreshing.get(userId);
}
