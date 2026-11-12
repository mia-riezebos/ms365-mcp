import { z, tool, enc, page, top, fileContent, decodeBase64, compact } from './common.mjs';

const location = {
  drive_id: z.string().optional().describe('Drive id (from list_site_drives or a shared item); omit for the signed-in user\'s OneDrive'),
  site_id: z.string().optional().describe('SharePoint site id (from search_sites); uses the site\'s default document library'),
};
const drivePath = ({ drive_id, site_id }) => (drive_id ? `/drives/${enc(drive_id)}` : site_id ? `/sites/${enc(site_id)}/drive` : '/me/drive');
const encPath = (p) => p.replace(/^\/+|\/+$/g, '').split('/').map(enc).join('/');
const itemRef = {
  item_id: z.string().optional(),
  path: z.string().optional().describe('Path from the drive root, e.g. Documents/report.docx'),
};
const itemPath = (a) => {
  const d = drivePath(a);
  if (a.item_id) return `${d}/items/${enc(a.item_id)}`;
  if (a.path && a.path.replace(/\//g, '')) return `${d}/root:/${encPath(a.path)}:`;
  return `${d}/root`;
};

const summarize = (i) => ({
  id: i.id,
  name: i.name,
  type: i.folder ? 'folder' : 'file',
  size: i.size,
  childCount: i.folder?.childCount,
  mimeType: i.file?.mimeType,
  modified: i.lastModifiedDateTime,
  modifiedBy: i.lastModifiedBy?.user?.displayName,
  path: i.parentReference?.path?.replace(/^\/drive(s\/[^/]+)?\/root:?/, '') || undefined,
  driveId: i.parentReference?.driveId,
  webUrl: i.webUrl,
});

export function registerFiles(server, graph) {
  tool(
    server,
    'list_files',
    { title: 'List files', description: 'Lists the files and folders in a OneDrive or SharePoint folder (the root if no item_id or path).', input: { ...location, ...itemRef, top: top(200, 100) } },
    async (a) => page(await graph('GET', `${itemPath(a)}/children`, { query: { $top: a.top } }), summarize),
  );

  tool(
    server,
    'search_files',
    { title: 'Search files', description: 'Searches a drive by file name and content.', input: { ...location, query: z.string(), top: top(100, 25) } },
    async (a) => page(await graph('GET', `${drivePath(a)}/root/search(q='${enc(a.query.replace(/'/g, "''"))}')`, { query: { $top: a.top } }), summarize),
  );

  tool(server, 'list_shared_with_me', { title: 'Files shared with me', description: 'Lists files and folders others have shared with the signed-in user.' }, async () =>
    page(await graph('GET', '/me/drive/sharedWithMe'), (i) => ({ ...summarize(i), driveId: i.remoteItem?.parentReference?.driveId, remoteId: i.remoteItem?.id })),
  );

  tool(server, 'get_file_info', { title: 'File details', description: 'Gets the metadata of a file or folder.', input: { ...location, ...itemRef } }, async (a) => summarize(await graph('GET', itemPath(a))));

  tool(
    server,
    'read_file',
    {
      title: 'Read a file',
      description: 'Downloads a file (up to 3 MB): text for text files, base64 otherwise. format=pdf or html converts Office documents first.',
      input: { ...location, ...itemRef, format: z.enum(['original', 'pdf', 'html']).default('original') },
    },
    async (a) => {
      const res = await graph('GET', `${itemPath(a)}/content`, { raw: true, query: { format: a.format !== 'original' ? a.format : undefined } });
      return fileContent(res, a.path || a.item_id);
    },
  );

  tool(
    server,
    'upload_file',
    {
      title: 'Upload a file',
      description: 'Creates or replaces a file of up to 4 MB from text or base64.',
      kind: 'write',
      input: {
        ...location,
        path: z.string().describe('Destination path from the drive root, including the file name'),
        text: z.string().optional(),
        content_base64: z.string().optional(),
        content_type: z.string().default('application/octet-stream'),
        conflict: z.enum(['replace', 'fail', 'rename']).default('fail'),
      },
    },
    async (a) => {
      if ((a.text === undefined) === (a.content_base64 === undefined)) throw new Error('Give exactly one of text or content_base64.');
      const body = a.text !== undefined ? new TextEncoder().encode(a.text) : decodeBase64(a.content_base64);
      const item = await graph('PUT', `${drivePath(a)}/root:/${encPath(a.path)}:/content`, {
        body,
        contentType: a.text !== undefined && a.content_type === 'application/octet-stream' ? 'text/plain' : a.content_type,
        query: { '@microsoft.graph.conflictBehavior': a.conflict },
      });
      return summarize(item);
    },
  );

  tool(
    server,
    'create_folder',
    { title: 'Create a folder', description: 'Creates a folder.', kind: 'write', input: { ...location, parent_path: z.string().default(''), parent_id: z.string().optional(), name: z.string() } },
    async (a) => summarize(await graph('POST', `${itemPath({ ...a, item_id: a.parent_id, path: a.parent_path })}/children`, { body: { name: a.name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' } })),
  );

  tool(
    server,
    'move_file',
    {
      title: 'Move or rename a file',
      description: 'Renames a file or folder and/or moves it to another folder in the same drive.',
      kind: 'write',
      input: { ...location, ...itemRef, new_name: z.string().optional(), destination_folder_id: z.string().optional() },
    },
    async (a) =>
      summarize(await graph('PATCH', itemPath(a), { body: compact({ name: a.new_name, parentReference: a.destination_folder_id ? { id: a.destination_folder_id } : undefined }) })),
  );

  tool(server, 'delete_file', { title: 'Delete a file', description: 'Moves a file or folder to the recycle bin.', kind: 'destructive', input: { ...location, ...itemRef } }, async (a) => {
    if (!a.item_id && !a.path) throw new Error('Give item_id or path.');
    await graph('DELETE', itemPath(a));
    return { deleted: true };
  });

  tool(server, 'search_sites', { title: 'Find SharePoint sites', description: 'Searches SharePoint sites by name.', input: { query: z.string() } }, async ({ query }) =>
    page(await graph('GET', '/sites', { query: { search: query } }), (s) => ({ id: s.id, name: s.displayName, webUrl: s.webUrl })),
  );

  tool(server, 'list_site_drives', { title: 'List document libraries', description: 'Lists the document libraries of a SharePoint site.', input: { site_id: z.string() } }, async ({ site_id }) =>
    page(await graph('GET', `/sites/${enc(site_id)}/drives`), (d) => ({ id: d.id, name: d.name, webUrl: d.webUrl })),
  );
}
