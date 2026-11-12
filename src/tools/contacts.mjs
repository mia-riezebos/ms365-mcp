import { userPath } from '../graph.mjs';
import { z, tool, enc, page, top, compact } from './common.mjs';

const owner = z.string().default('me').describe('Whose contacts: "me" or the email address of a shared mailbox');
const FIELDS = 'id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,homePhones,companyName,jobTitle,department,businessAddress,homeAddress,birthday,personalNotes,categories';

const contactFields = {
  given_name: z.string().optional(),
  surname: z.string().optional(),
  display_name: z.string().optional(),
  emails: z.array(z.string().email()).optional(),
  mobile_phone: z.string().optional(),
  business_phones: z.array(z.string()).optional(),
  company: z.string().optional(),
  job_title: z.string().optional(),
  notes: z.string().optional(),
  categories: z.array(z.string()).optional(),
};
const contactBody = (a) =>
  compact({
    givenName: a.given_name,
    surname: a.surname,
    displayName: a.display_name,
    emailAddresses: a.emails?.map((address) => ({ address })),
    mobilePhone: a.mobile_phone,
    businessPhones: a.business_phones,
    companyName: a.company,
    jobTitle: a.job_title,
    personalNotes: a.notes,
    categories: a.categories,
  });

export function registerContacts(server, graph) {
  tool(
    server,
    'list_contacts',
    {
      title: 'List or find contacts',
      description: 'Lists Outlook contacts, or finds them by name or email.',
      input: { owner, search: z.string().optional().describe('Start of a name, or a full email address'), folder_id: z.string().optional(), top: top(100, 50) },
    },
    async ({ owner: o, search, folder_id, top }) => {
      const base = `${userPath(o)}${folder_id ? `/contactFolders/${enc(folder_id)}` : ''}/contacts`;
      const q = search?.replace(/'/g, "''");
      const $filter = q && (q.includes('@') ? `emailAddresses/any(a:a/address eq '${q}')` : ['displayName', 'givenName', 'surname', 'companyName'].map((f) => `startswith(${f},'${q}')`).join(' or '));
      return page(await graph('GET', base, { query: { $top: top, $select: FIELDS, $filter, $orderby: $filter ? undefined : 'displayName' } }));
    },
  );

  tool(server, 'list_contact_folders', { title: 'List contact folders', description: 'Lists contact folders.', input: { owner } }, async ({ owner: o }) =>
    page(await graph('GET', `${userPath(o)}/contactFolders`)),
  );

  tool(server, 'get_contact', { title: 'Read a contact', description: 'Gets one contact.', input: { owner, contact_id: z.string() } }, ({ owner: o, contact_id }) =>
    graph('GET', `${userPath(o)}/contacts/${enc(contact_id)}`),
  );

  tool(
    server,
    'create_contact',
    { title: 'Create a contact', description: 'Adds an Outlook contact.', kind: 'write', input: { owner, folder_id: z.string().optional(), ...contactFields } },
    (a) => graph('POST', `${userPath(a.owner)}${a.folder_id ? `/contactFolders/${enc(a.folder_id)}` : ''}/contacts`, { body: contactBody(a) }),
  );

  tool(
    server,
    'update_contact',
    { title: 'Update a contact', description: 'Changes the given fields of a contact; lists replace the current ones.', kind: 'write', input: { owner, contact_id: z.string(), ...contactFields } },
    (a) => graph('PATCH', `${userPath(a.owner)}/contacts/${enc(a.contact_id)}`, { body: contactBody(a) }),
  );

  tool(server, 'delete_contact', { title: 'Delete a contact', description: 'Deletes a contact.', kind: 'destructive', input: { owner, contact_id: z.string() } }, async ({ owner: o, contact_id }) => {
    await graph('DELETE', `${userPath(o)}/contacts/${enc(contact_id)}`);
    return { deleted: true };
  });
}
