import { z } from 'zod';

export { z };
export const enc = encodeURIComponent;

const ANNOTATIONS = {
  read: { readOnlyHint: true, openWorldHint: true },
  write: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
  // Sends something to other people, or deletes.
  destructive: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
};

export const ok = (data) => ({ content: [{ type: 'text', text: typeof data === 'string' ? data : JSON.stringify(data, null, 2) }] });

export function tool(server, name, { title, description, input = {}, kind = 'read' }, handler) {
  server.registerTool(name, { title, description, inputSchema: input, annotations: ANNOTATIONS[kind] }, async (args) => {
    try {
      return ok(await handler(args));
    } catch (err) {
      return { isError: true, content: [{ type: 'text', text: err.message }] };
    }
  });
}

export const mailbox = z
  .string()
  .default('me')
  .describe('Mailbox to act on: "me" for the signed-in user, or the email address of a shared or delegated mailbox (e.g. info@example.com)');

export const top = (max = 50, def = 25) => z.number().int().min(1).max(max).default(def);
export const timeZone = z.string().default('Europe/Amsterdam').describe('IANA or Windows time zone for returned and given times');
export const dateTime = (what) => z.string().describe(`${what}, ISO 8601 local time without offset, e.g. 2026-11-12T09:30:00`);

// Graph list response -> { count, items, nextLink }; pass nextLink to graph_request to get the next page.
export const page = (res, map = (x) => x) => ({ count: res.value.length, items: res.value.map(map), nextLink: res['@odata.nextLink'] });

export const addr = (r) => (r?.emailAddress ? `${r.emailAddress.name || ''} <${r.emailAddress.address}>`.trim() : undefined);
export const toRecipients = (list = []) => list.map((address) => ({ emailAddress: { address } }));
export const recipients = z.array(z.string().email()).describe('Email addresses');

export const bodyInput = z.string().optional();
export const bodyType = z.enum(['text', 'html']).default('text');
export const itemBody = (content, type) => (content === undefined ? undefined : { contentType: type === 'html' ? 'HTML' : 'Text', content });

const textual = (ct = '') => /^text\/|json|xml|csv|javascript|yaml/.test(ct);
export async function fileContent(res, name, maxBytes = 3 * 1024 * 1024) {
  const contentType = res.headers.get('content-type') || '';
  const bytes = new Uint8Array(await res.arrayBuffer());
  if (bytes.length > maxBytes) return { name, contentType, size: bytes.length, note: `Larger than ${maxBytes} bytes; not returned.` };
  if (textual(contentType)) return { name, contentType, size: bytes.length, text: new TextDecoder().decode(bytes) };
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return { name, contentType, size: bytes.length, contentBase64: btoa(bin) };
}

export const decodeBase64 = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));

// Drops undefined values so PATCH bodies only carry the fields given.
export const compact = (obj) => Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== undefined));
