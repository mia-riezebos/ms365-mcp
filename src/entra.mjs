// Delegated Microsoft Entra sign-in and per-connection token storage.
// Each MCP connection has its own record in KV (encrypted with TOKEN_ENC_KEY): the Entra account, the scopes
// granted to that connection, and its refresh token. Access tokens are requested per API with exactly the
// connection's scopes for that API, and the rotated refresh token is stored on every refresh.
import { BASE_SCOPES, GRAPH_ORIGIN, OIDC_SCOPES, entraScope, isGraphScope, normalizeScope, resourceOf, union } from './scopes.mjs';

const authority = (env) => `https://login.microsoftonline.com/${env.MS365_TENANT_ID || 'organizations'}/oauth2/v2.0`;
export const redirectUri = (req) => new URL('/auth/callback', req.url).toString();
const scopeParam = (scopes) => union(OIDC_SCOPES, scopes).map(entraScope).join(' ');

const b64url = (bytes) => btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const fromB64url = (s) => Uint8Array.from(atob(s.replace(/-/g, '+').replace(/_/g, '/')), (c) => c.charCodeAt(0));
export const randomId = () => b64url(crypto.getRandomValues(new Uint8Array(24)));

export const pkceChallenge = async (verifier) => b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));

export async function pkcePair() {
  const verifier = randomId();
  return { verifier, challenge: await pkceChallenge(verifier) };
}

// Microsoft sign-in URL asking for scopes (plus OpenID Connect and offline_access).
export function authorizeUrl(env, req, { state, codeChallenge, scopes, loginHint }) {
  const url = new URL(`${authority(env)}/authorize`);
  url.search = new URLSearchParams({
    client_id: env.MS365_APP_CLIENT_ID,
    response_type: 'code',
    response_mode: 'query',
    redirect_uri: redirectUri(req),
    scope: scopeParam(scopes),
    state,
    code_challenge: codeChallenge,
    code_challenge_method: 'S256',
    ...(loginHint ? { login_hint: loginHint } : { prompt: 'select_account' }),
  }).toString();
  return url.toString();
}

export class EntraError extends Error {
  constructor(body) {
    super(`Microsoft sign-in error: ${body.error}: ${body.error_description?.split('\r\n')[0]}`);
    this.code = body.error;
  }
}

async function tokenRequest(env, scopes, params) {
  const res = await fetch(`${authority(env)}/token`, {
    method: 'POST',
    body: new URLSearchParams({ client_id: env.MS365_APP_CLIENT_ID, client_secret: env.MS365_APP_CLIENT_SECRET, scope: scopeParam(scopes), ...params }),
  });
  const body = await res.json();
  if (!res.ok) throw new EntraError(body);
  return body;
}

const graphScopesOf = (scopes) => {
  const graph = scopes.filter(isGraphScope);
  return graph.length ? graph : BASE_SCOPES;
};

// Redeems a sign-in code for Graph tokens covering the Graph scopes in scopes.
export const redeemCode = (env, req, code, codeVerifier, scopes) =>
  tokenRequest(env, graphScopesOf(scopes), { grant_type: 'authorization_code', code, redirect_uri: redirectUri(req), code_verifier: codeVerifier });

export function accountOf(tokens) {
  const claims = JSON.parse(new TextDecoder().decode(fromB64url(tokens.id_token.split('.')[1])));
  return { userId: claims.oid, tenantId: claims.tid, upn: (claims.preferred_username || claims.upn || '').toLowerCase(), name: claims.name };
}

// Scopes Microsoft actually granted, of the ones asked for.
const grantedOf = (tokens, asked) => {
  const got = new Set((tokens.scope || '').split(' ').map(normalizeScope).filter(Boolean).map((s) => s.toLowerCase()));
  return asked.filter((s) => got.has(s.toLowerCase()));
};

let keyPromise;
const encKey = (env) => (keyPromise ??= crypto.subtle.importKey('raw', fromB64url(env.TOKEN_ENC_KEY), 'AES-GCM', false, ['encrypt', 'decrypt']));

async function seal(env, value) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const data = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await encKey(env), new TextEncoder().encode(JSON.stringify(value)));
  return `${b64url(iv)}.${b64url(data)}`;
}

async function unseal(env, sealed) {
  const [iv, data] = sealed.split('.').map(fromB64url);
  return JSON.parse(new TextDecoder().decode(await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, await encKey(env), data)));
}

const connKey = (connId) => `conn:${connId}`;

async function saveConnection(env, connId, conn) {
  await env.OAUTH_KV.put(connKey(connId), JSON.stringify({ upn: conn.upn, scopes: conn.scopes, updatedAt: new Date().toISOString(), sealed: await seal(env, conn) }));
}

export async function loadConnection(env, connId) {
  const raw = await env.OAUTH_KV.get(connKey(connId));
  return raw ? unseal(env, JSON.parse(raw).sealed) : null;
}

const cached = (tokens) => ({ token: tokens.access_token, expiresAt: Date.now() + (tokens.expires_in - 120) * 1000 });

// Creates a connection from a first sign-in: identity only, with BASE_SCOPES.
export async function createConnection(env, account, tokens) {
  const connId = randomId();
  const scopes = union(BASE_SCOPES, grantedOf(tokens, BASE_SCOPES));
  await saveConnection(env, connId, { ...account, scopes, refreshToken: tokens.refresh_token, access: { [GRAPH_ORIGIN]: cached(tokens) }, createdAt: new Date().toISOString() });
  return connId;
}

// Adds the scopes from a request_scopes sign-in. Graph scopes count as granted if the Graph token has them;
// each other API is checked by getting a token for it. Returns { granted, refused }.
export async function addScopes(env, connId, conn, tokens, requested) {
  conn.refreshToken = tokens.refresh_token || conn.refreshToken;
  const graphAsked = requested.filter(isGraphScope);
  const granted = grantedOf(tokens, graphAsked);
  conn.access = { [GRAPH_ORIGIN]: cached(tokens) };
  for (const resource of new Set(requested.filter((s) => !isGraphScope(s)).map(resourceOf))) {
    const asked = requested.filter((s) => !isGraphScope(s) && resourceOf(s) === resource);
    try {
      const t = await tokenRequest(env, asked, { grant_type: 'refresh_token', refresh_token: conn.refreshToken });
      if (t.refresh_token) conn.refreshToken = t.refresh_token;
      conn.access[resource] = cached(t);
      granted.push(...grantedOf(t, asked));
    } catch {}
  }
  conn.scopes = union(conn.scopes, granted);
  await saveConnection(env, connId, conn);
  const have = new Set(conn.scopes.map((s) => s.toLowerCase()));
  return { granted: conn.scopes, refused: requested.filter((s) => !have.has(s.toLowerCase())) };
}

const refreshing = new Map();

// A valid access token for one API (Graph by default), covering this connection's scopes for it.
export async function accessToken(env, connId, resource = GRAPH_ORIGIN) {
  const conn = await loadConnection(env, connId);
  if (!conn) throw new Error('This connection has no Microsoft account; reconnect the MCP server.');
  const hit = conn.access?.[resource];
  if (hit && hit.expiresAt > Date.now()) return hit.token;
  const scopes = resource === GRAPH_ORIGIN ? graphScopesOf(conn.scopes) : conn.scopes.filter((s) => !isGraphScope(s) && resourceOf(s) === resource);
  if (!scopes.length) throw new Error(`This connection has no scopes for ${resource}; call request_scopes first.`);
  const key = `${connId} ${resource}`;
  if (!refreshing.has(key)) {
    refreshing.set(
      key,
      tokenRequest(env, scopes, { grant_type: 'refresh_token', refresh_token: conn.refreshToken })
        .then(async (tokens) => {
          conn.access = { ...conn.access, [resource]: cached(tokens) };
          if (tokens.refresh_token) conn.refreshToken = tokens.refresh_token;
          await saveConnection(env, connId, conn);
          return tokens.access_token;
        })
        .catch((err) => {
          if (err.code === 'invalid_grant') throw new Error('Microsoft sign-in has expired or was revoked; reconnect the MCP server.');
          throw err;
        })
        .finally(() => refreshing.delete(key)),
    );
  }
  return refreshing.get(key);
}

// Scope requests: a one-time link that sends the user to Microsoft for more scopes on one connection.
// The PKCE verifier is fixed when the link is made, so opening the link again (a link preview, an
// in-app browser handing off to Safari) sends the same challenge instead of invalidating the first.
const requestKey = (id) => `scope-request:${id}`;
export async function createScopeRequest(env, connId, scopes, reason) {
  const id = randomId();
  await env.OAUTH_KV.put(requestKey(id), JSON.stringify({ connId, scopes, reason, verifier: randomId() }), { expirationTtl: 60 * 60 });
  return id;
}
export async function loadScopeRequest(env, id) {
  const raw = await env.OAUTH_KV.get(requestKey(id));
  return raw ? JSON.parse(raw) : null;
}
export const deleteScopeRequest = (env, id) => env.OAUTH_KV.delete(requestKey(id));
