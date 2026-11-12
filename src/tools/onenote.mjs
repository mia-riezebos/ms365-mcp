import { z, tool, enc, page, top } from './common.mjs';

const esc = (s) => s.replace(/[&<>"]/g, (c) => `&#${c.charCodeAt(0)};`);

export function registerOneNote(server, graph) {
  tool(server, 'list_notebooks', { title: 'List OneNote notebooks', description: 'Lists the signed-in user\'s OneNote notebooks.' }, async () =>
    page(await graph('GET', '/me/onenote/notebooks', { query: { $select: 'id,displayName,lastModifiedDateTime,links' } }), (n) => ({ id: n.id, name: n.displayName, modified: n.lastModifiedDateTime, webUrl: n.links?.oneNoteWebUrl?.href })),
  );

  tool(server, 'list_onenote_sections', { title: 'List OneNote sections', description: 'Lists sections, of one notebook or all.', input: { notebook_id: z.string().optional() } }, async ({ notebook_id }) =>
    page(await graph('GET', notebook_id ? `/me/onenote/notebooks/${enc(notebook_id)}/sections` : '/me/onenote/sections', { query: { $select: 'id,displayName,parentNotebook' } }), (s) => ({ id: s.id, name: s.displayName, notebook: s.parentNotebook?.displayName })),
  );

  tool(
    server,
    'list_onenote_pages',
    { title: 'List OneNote pages', description: 'Lists pages, of one section or recent pages across notebooks.', input: { section_id: z.string().optional(), top: top(100, 25) } },
    async ({ section_id, top }) =>
      page(
        await graph('GET', section_id ? `/me/onenote/sections/${enc(section_id)}/pages` : '/me/onenote/pages', { query: { $top: top, $select: 'id,title,createdDateTime,lastModifiedDateTime,parentSection,links', $orderby: 'lastModifiedDateTime desc' } }),
        (p) => ({ id: p.id, title: p.title, modified: p.lastModifiedDateTime, section: p.parentSection?.displayName, webUrl: p.links?.oneNoteWebUrl?.href }),
      ),
  );

  tool(server, 'read_onenote_page', { title: 'Read a OneNote page', description: 'Gets the HTML content of a OneNote page.', input: { page_id: z.string() } }, async ({ page_id }) => ({
    html: await graph('GET', `/me/onenote/pages/${enc(page_id)}/content`),
  }));

  tool(
    server,
    'create_onenote_page',
    { title: 'Create a OneNote page', description: 'Adds a page to a section.', kind: 'write', input: { section_id: z.string(), title: z.string(), html_body: z.string().describe('Page body as HTML, e.g. <p>Notes</p>') } },
    async ({ section_id, title, html_body }) => {
      const doc = `<!DOCTYPE html><html><head><title>${esc(title)}</title></head><body>${html_body}</body></html>`;
      const p = await graph('POST', `/me/onenote/sections/${enc(section_id)}/pages`, { body: doc, contentType: 'text/html' });
      return { id: p.id, title: p.title, webUrl: p.links?.oneNoteWebUrl?.href };
    },
  );
}
