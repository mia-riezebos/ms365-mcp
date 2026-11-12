import { OAuthProvider, AuthorizationError, CimdFetchError, authorizationErrorRedirect } from '@cloudflare/workers-oauth-provider';
import { WebStandardStreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js';
import { createMcpServer } from './tools/index.mjs';
import { graphClient } from './graph.mjs';
import { authorizeUrl, graphAccessToken, pkcePair, redeemCode, storeUserTokens } from './entra.mjs';

const DAY = 24 * 60 * 60;
const escape = (v) => String(v).replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);
const html = (body, init = {}) => {
  const headers = new Headers(init.headers);
  headers.set('Content-Type', 'text/html; charset=utf-8');
  return new Response(`<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<style>body{font:16px system-ui,sans-serif;max-width:40rem;margin:3rem auto;padding:0 1rem;color:#222}code{background:#eee;padding:0 .3em}button{font:inherit;padding:.4em 1em;margin-right:.5em}</style>${body}`, { ...init, headers });
};
const text = (msg, status = 400) => new Response(msg, { status, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
const allowedUsers = (env) => (env.ALLOWED_USERS || '').split(',').map((s) => s.trim().toLowerCase()).filter(Boolean);
const LOOPBACK = ['localhost', '127.0.0.1', '[::1]'];
// Any HTTPS, loopback or app-scheme redirect unless ALLOWED_REDIRECT_HOSTS narrows it; the consent page names the host either way.
function redirectAllowed(env, redirectUri) {
  const url = new URL(redirectUri);
  const hosts = (env.ALLOWED_REDIRECT_HOSTS || '').split(',').map((s) => s.trim()).filter(Boolean);
  if (hosts.length) return hosts.includes(url.hostname);
  if (url.protocol === 'https:') return true;
  if (url.protocol === 'http:') return LOOPBACK.includes(url.hostname);
  return !['javascript:', 'data:', 'file:', 'blob:'].includes(url.protocol);
}

const mcpHandler = {
  async fetch(request, env, ctx) {
    if (request.method !== 'POST') return text('Method not allowed', 405);
    const { userId } = ctx.props;
    const server = createMcpServer(graphClient(() => graphAccessToken(env, userId)));
    const transport = new WebStandardStreamableHTTPServerTransport({ sessionIdGenerator: undefined, enableJsonResponse: true });
    await server.connect(transport);
    return transport.handleRequest(request);
  },
};

function consentPage(details, handle) {
  return html(`<title>Connect ${escape(details.clientName)}</title>
<h1>Connect ${escape(details.clientName)} to Microsoft 365?</h1>
<p>${details.clientDomain ? `Published by <strong>${escape(details.clientDomain)}</strong>.` : 'This app registered itself; its name is not verified.'}
Access will be sent to <strong>${escape(details.redirectHost)}</strong>.</p>
${details.redirectIsLoopback ? '<p><strong>This sends access to an app on your computer.</strong> Continue only if you just started connecting from it.</p>' : ''}
<p>It will act as you in Microsoft 365: read, send and organize mail (including shared mailboxes you can open), calendars, contacts, Teams chats and channels, Planner, To Do, OneDrive and SharePoint files, and OneNote. Allow takes you to Microsoft sign-in.</p>
<form method="post"><input type="hidden" name="handle" value="${escape(handle)}">
<button name="decision" value="approve">Allow</button><button name="decision" value="deny">Deny</button></form>`);
}

async function authorize(request, env) {
  const oauth = env.OAUTH_PROVIDER;
  if (request.method === 'GET') {
    const authRequest = await oauth.parseAuthRequest(request);
    if (!redirectAllowed(env, authRequest.redirectUri)) return text(`Redirect to ${new URL(authRequest.redirectUri).origin} is not allowed.`);
    const details = await oauth.describeConsent(authRequest);
    const consent = await oauth.beginConsent(authRequest);
    return withHeaders(consentPage(details, consent.handle), consent.headers);
  }
  if (request.method !== 'POST') return text('Method not allowed', 405);
  const form = await request.formData();
  const handle = String(form.get('handle'));
  if (form.get('decision') !== 'approve') {
    const denied = await oauth.denyConsent(request, handle);
    return new Response(null, { status: 302, headers: denied.headers });
  }
  const approved = await oauth.approveConsent(request, handle);
  const { verifier, challenge } = await pkcePair();
  const { state, headers } = await oauth.beginUpstream(approved.request, { data: { verifier }, headers: approved.headers });
  headers.set('Location', authorizeUrl(env, request, { state, codeChallenge: challenge }));
  return new Response(null, { status: 302, headers });
}

function withHeaders(response, headers) {
  for (const [k, v] of headers) response.headers.set(k, v);
  return response;
}

async function callback(request, env) {
  const oauth = env.OAUTH_PROVIDER;
  const params = new URL(request.url).searchParams;
  const { request: original, data, headers } = await oauth.finishUpstream(request);
  if (params.get('error')) {
    headers.set('Location', authorizationErrorRedirect(original, 'access_denied'));
    return new Response(null, { status: 302, headers });
  }
  const tokens = await redeemCode(env, request, params.get('code'), data.verifier);
  const allowed = allowedUsers(env);
  const user = await storeUserTokens(env, tokens, ({ tenantId, upn }) => {
    if (tenantId !== env.MS365_TENANT_ID) return 'This Microsoft account is not in the allowed tenant.';
    if (allowed.length && !allowed.includes(upn)) return `${upn} is not allowed to use this server.`;
  });
  if (user.refusal) return text(user.refusal, 403);
  const { redirectTo } = await oauth.completeAuthorization({
    request: original,
    userId: user.userId,
    metadata: { upn: user.upn },
    scope: original.scope,
    props: { userId: user.userId, upn: user.upn },
  });
  headers.set('Location', redirectTo);
  return new Response(null, { status: 302, headers });
}

const defaultHandler = {
  async fetch(request, env) {
    const { pathname, origin } = new URL(request.url);
    try {
      if (pathname === '/authorize') return await authorize(request, env);
      if (pathname === '/auth/callback') return await callback(request, env);
    } catch (error) {
      if (error instanceof AuthorizationError && error.redirectTo) return Response.redirect(error.redirectTo, 302);
      if (error instanceof AuthorizationError) return text(error.description);
      if (error instanceof CimdFetchError) return text('This app could not be verified.');
      throw error;
    }
    if (pathname === '/healthz') return text('ok', 200);
    if (pathname === '/')
      return html(`<title>Microsoft 365 MCP</title><h1>Microsoft 365 MCP</h1>
<p>Remote MCP server for Microsoft 365 as the signed-in user: mail (incl. shared mailboxes), calendar, contacts, Teams, Planner, To Do, OneDrive/SharePoint and OneNote, plus a raw Graph call.</p>
<p>MCP endpoint: <code>${escape(origin)}/mcp</code>. Connect it as an OAuth MCP server; connecting signs you in with Microsoft.</p>`);
    return text('Not found', 404);
  },
};

const providerOptions = (env) => ({
  apiRoute: '/mcp',
  apiHandler: mcpHandler,
  defaultHandler,
  authorizeEndpoint: '/authorize',
  tokenEndpoint: '/oauth/token',
  clientRegistrationEndpoint: '/oauth/register',
  scopesSupported: ['mail'],
  accessTokenTTL: 60 * 60,
  // Grants live while the client keeps refreshing, matching Entra's 90-day sliding refresh token window.
  refreshTokenTTL: 90 * DAY,
  refreshTokenIdleTTL: 90 * DAY,
  clientRegistrationTTL: undefined,
  resourceMetadata: { resource: `${env.PUBLIC_BASE_URL}/mcp`, resource_name: 'Microsoft 365' },
});

// The resource URL comes from env, so the provider is built on the first request.
let provider;
export default {
  fetch(request, env, ctx) {
    provider ??= new OAuthProvider(providerOptions(env));
    return provider.fetch(request, env, ctx);
  },
};
