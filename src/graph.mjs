const ORIGIN = 'https://graph.microsoft.com';

export class GraphError extends Error {}

// "me" or empty for the signed-in user, otherwise /users/{address or id} (shared and delegated mailboxes, other people's calendars).
export function userPath(user) {
  if (!user || user.toLowerCase() === 'me') return '/me';
  return `/users/${encodeURIComponent(user)}`;
}
export const mailboxPath = userPath;

const isRawBody = (b) => typeof b === 'string' || b instanceof Uint8Array || b instanceof ArrayBuffer;

// Returns graph(method, path, options) that calls Microsoft Graph with a token from getToken().
// path is relative to the API version, or a full https://graph.microsoft.com URL such as an @odata.nextLink.
// options: query, body (JSON, or a string/bytes with contentType), headers, version ('v1.0' | 'beta'), raw (return the Response).
export function graphClient(getToken) {
  return async function graph(method, path, { query, body, headers, contentType, version = 'v1.0', raw = false } = {}) {
    const url = new URL(path.startsWith('https://') ? path : `${ORIGIN}/${version}${path.startsWith('/') ? '' : '/'}${path}`);
    if (url.origin !== ORIGIN) throw new GraphError(`Only ${ORIGIN} URLs can be called.`);
    for (const [k, v] of Object.entries(query || {})) if (v !== undefined && v !== null && v !== '') url.searchParams.set(k, String(v));
    const hasBody = body !== undefined && body !== null;
    const res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${await getToken()}`,
        ...(hasBody ? { 'Content-Type': contentType || (isRawBody(body) ? 'application/octet-stream' : 'application/json') } : {}),
        ...headers,
      },
      body: hasBody ? (isRawBody(body) ? body : JSON.stringify(body)) : undefined,
    });
    if (!res.ok) {
      const text = await res.text();
      let err;
      try {
        err = JSON.parse(text).error;
      } catch {}
      throw new GraphError(`Graph ${res.status} ${err?.code || ''}: ${err?.message || text.slice(0, 500)}`);
    }
    if (raw) return res;
    if (res.status === 202 || res.status === 204) return null;
    const text = await res.text();
    if (!text) return null;
    return (res.headers.get('content-type') || '').includes('json') ? JSON.parse(text) : text;
  };
}
