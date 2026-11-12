// Delegated permission scopes: which tools each one unlocks, and checks for raw Graph calls.
// Graph scopes are kept in their short form (Mail.Read); scopes of other APIs as full URIs
// (https://outlook.office365.com/EWS.AccessAsUser.All).

export const GRAPH_ORIGIN = 'https://graph.microsoft.com';
export const OIDC_SCOPES = ['openid', 'profile', 'email', 'offline_access'];
// Granted on first sign-in: enough to identify the account.
export const BASE_SCOPES = ['User.Read'];

const shared = (...names) => names.flatMap((n) => [n, `${n}.Shared`]);
const MAIL_READ = shared('Mail.Read', 'Mail.ReadWrite');
const MAIL_WRITE = shared('Mail.ReadWrite');
const MAIL_SEND = shared('Mail.Send');
const CAL_READ = shared('Calendars.Read', 'Calendars.ReadWrite');
const CAL_WRITE = shared('Calendars.ReadWrite');
const CONTACTS_READ = shared('Contacts.Read', 'Contacts.ReadWrite');
const CONTACTS_WRITE = shared('Contacts.ReadWrite');
const TASKS_READ = ['Tasks.Read', 'Tasks.ReadWrite'];
const TASKS_WRITE = ['Tasks.ReadWrite'];
const PLANNER_READ = [...TASKS_READ, 'Group.Read.All', 'Group.ReadWrite.All'];
const PLANNER_WRITE = [...TASKS_WRITE, 'Group.ReadWrite.All'];
const SITES_READ = ['Sites.Read.All', 'Sites.ReadWrite.All'];
const FILES_READ = ['Files.Read', 'Files.ReadWrite', 'Files.Read.All', 'Files.ReadWrite.All', ...SITES_READ];
const FILES_WRITE = ['Files.ReadWrite', 'Files.ReadWrite.All', 'Sites.ReadWrite.All'];
const NOTES_READ = ['Notes.Read', 'Notes.ReadWrite', 'Notes.Read.All', 'Notes.ReadWrite.All'];
const NOTES_CREATE = ['Notes.Create', 'Notes.ReadWrite', 'Notes.ReadWrite.All'];
const CHAT_READ = ['Chat.Read', 'Chat.ReadWrite'];
const ALWAYS = [];

// Tool name -> scopes, any one of which makes the tool available. ALWAYS: no scope needed.
export const TOOL_SCOPES = {
  whoami: ALWAYS,
  request_scopes: ALWAYS,
  graph_request: ALWAYS,
  list_mail_folders: MAIL_READ,
  list_messages: MAIL_READ,
  get_message: MAIL_READ,
  list_attachments: MAIL_READ,
  get_attachment: MAIL_READ,
  create_draft: MAIL_WRITE,
  send_mail: MAIL_SEND,
  send_draft: MAIL_SEND,
  reply_message: MAIL_SEND,
  forward_message: MAIL_SEND,
  move_message: MAIL_WRITE,
  update_message: MAIL_WRITE,
  list_calendars: CAL_READ,
  list_events: CAL_READ,
  get_event: CAL_READ,
  create_event: CAL_WRITE,
  update_event: CAL_WRITE,
  delete_event: CAL_WRITE,
  respond_to_event: CAL_WRITE,
  get_free_busy: ['Calendars.ReadBasic', ...CAL_READ],
  list_contacts: CONTACTS_READ,
  list_contact_folders: CONTACTS_READ,
  get_contact: CONTACTS_READ,
  create_contact: CONTACTS_WRITE,
  update_contact: CONTACTS_WRITE,
  delete_contact: CONTACTS_WRITE,
  list_teams: ['Team.ReadBasic.All', 'TeamSettings.Read.All', 'TeamSettings.ReadWrite.All'],
  list_channels: ['Channel.ReadBasic.All', 'ChannelSettings.Read.All', 'ChannelSettings.ReadWrite.All'],
  list_channel_messages: ['ChannelMessage.Read.All'],
  send_channel_message: ['ChannelMessage.Send'],
  list_chats: ['Chat.ReadBasic', ...CHAT_READ],
  list_chat_messages: CHAT_READ,
  send_chat_message: ['ChatMessage.Send', 'Chat.ReadWrite'],
  list_plans: PLANNER_READ,
  list_buckets: PLANNER_READ,
  list_planner_tasks: PLANNER_READ,
  get_planner_task: PLANNER_READ,
  create_planner_task: PLANNER_WRITE,
  update_planner_task: PLANNER_WRITE,
  delete_planner_task: PLANNER_WRITE,
  list_todo_lists: TASKS_READ,
  list_todo_tasks: TASKS_READ,
  create_todo_task: TASKS_WRITE,
  update_todo_task: TASKS_WRITE,
  delete_todo_task: TASKS_WRITE,
  list_files: FILES_READ,
  search_files: FILES_READ,
  list_shared_with_me: ['Files.Read.All', 'Files.ReadWrite.All'],
  get_file_info: FILES_READ,
  read_file: FILES_READ,
  upload_file: FILES_WRITE,
  create_folder: FILES_WRITE,
  move_file: FILES_WRITE,
  delete_file: FILES_WRITE,
  search_sites: SITES_READ,
  list_site_drives: SITES_READ,
  list_notebooks: NOTES_READ,
  list_onenote_sections: NOTES_READ,
  list_onenote_pages: NOTES_READ,
  read_onenote_page: NOTES_READ,
  create_onenote_page: NOTES_CREATE,
};

export const SCOPE_GUIDE = `Common Microsoft Graph scopes (pass the short names):
- Mail: Mail.Read (read), Mail.ReadWrite (drafts, move, flag), Mail.Send (send, reply, forward). Add .Shared (Mail.Read.Shared, Mail.ReadWrite.Shared, Mail.Send.Shared) to use shared or delegated mailboxes. MailboxSettings.ReadWrite for rules, auto-replies and settings via graph_request.
- Calendar: Calendars.Read, Calendars.ReadWrite (and .Shared variants for other people's or shared calendars).
- Contacts: Contacts.Read, Contacts.ReadWrite (and .Shared).
- Teams: Team.ReadBasic.All, Channel.ReadBasic.All, ChannelMessage.Read.All, ChannelMessage.Send, Chat.Read, Chat.ReadWrite, ChatMessage.Send.
- Planner and To Do: Tasks.Read, Tasks.ReadWrite (Group.Read.All for plans of groups you are not in).
- Files: Files.Read, Files.ReadWrite (own OneDrive), Files.Read.All / Files.ReadWrite.All (shared files), Sites.Read.All / Sites.ReadWrite.All (SharePoint).
- OneNote: Notes.Read, Notes.ReadWrite, Notes.Create.
- Anything else in Graph (People.Read, Presence.Read, OnlineMeetings.ReadWrite, Group.Read.All, ...) works through graph_request.
Scopes of other APIs go in as full URIs, e.g. https://outlook.office365.com/EWS.AccessAsUser.All, and are called with api_request.`;

const OIDC = new Set(OIDC_SCOPES);
const GRAPH_PREFIX = `${GRAPH_ORIGIN}/`;

// "https://graph.microsoft.com/Mail.Read" -> "Mail.Read"; other API scopes stay full URIs. Returns null if malformed.
export function normalizeScope(raw) {
  const s = String(raw).trim();
  if (s.toLowerCase().startsWith(GRAPH_PREFIX)) return normalizeScope(s.slice(GRAPH_PREFIX.length));
  if (/^[A-Za-z][\w.-]*$/.test(s)) return s;
  if (/^(https:\/\/[\w.-]+(:\d+)?|api:\/\/[\w.-]+)\/[\w.\-/]+$/.test(s)) return s;
  return null;
}

export const isGraphScope = (s) => !s.includes('://');
export const isOidcScope = (s) => OIDC.has(s.toLowerCase());
// The API a scope belongs to: Graph, or the scope URI without its last segment.
export const resourceOf = (s) => (isGraphScope(s) ? GRAPH_ORIGIN : s.slice(0, s.lastIndexOf('/')));
// Form Entra expects in a scope parameter.
export const entraScope = (s) => (isOidcScope(s) || !isGraphScope(s) ? s : `${GRAPH_PREFIX}${s}`);

export function hasAny(granted, needed) {
  const have = new Set(granted.map((s) => s.toLowerCase()));
  return needed.some((s) => have.has(s.toLowerCase()));
}

export function union(...lists) {
  const seen = new Map();
  for (const s of lists.flat()) if (!seen.has(s.toLowerCase())) seen.set(s.toLowerCase(), s);
  return [...seen.values()];
}

// Unknown tools stay hidden.
export function toolAllowed(name, granted) {
  const needed = TOOL_SCOPES[name];
  if (!needed) return name === 'api_request' && granted.some((s) => !isGraphScope(s));
  return needed.length === 0 || hasAny(granted, needed);
}

const readOnly = (scopes) => scopes;
const writable = (scopes) => scopes.filter((s) => /ReadWrite|Send|Create/i.test(s));
// Raw Graph calls into these areas need one of the area's scopes on this connection (a write scope unless the call only reads).
const AREAS = [
  { area: 'mail', match: /^(messages|mailfolders|sendmail)\b/, read: [...MAIL_READ, ...MAIL_SEND], write: [...MAIL_WRITE, ...MAIL_SEND] },
  { area: 'mailbox settings', match: /^mailboxsettings\b/, read: ['MailboxSettings.Read', 'MailboxSettings.ReadWrite'], write: ['MailboxSettings.ReadWrite'] },
  { area: 'calendar', match: /^(events|calendar|calendars|calendarview|calendargroups|findmeetingtimes)\b/, read: ['Calendars.ReadBasic', ...CAL_READ], write: CAL_WRITE },
  { area: 'contacts', match: /^(contacts|contactfolders)\b/, read: CONTACTS_READ, write: CONTACTS_WRITE },
  { area: 'files', match: /^(drive|drives)\b/, read: FILES_READ, write: FILES_WRITE },
  { area: 'SharePoint', match: /^sites\b/, read: [...SITES_READ, 'Files.Read.All', 'Files.ReadWrite.All'], write: ['Sites.ReadWrite.All', 'Files.ReadWrite.All'] },
  { area: 'Planner', match: /^planner\b/, read: PLANNER_READ, write: PLANNER_WRITE },
  { area: 'To Do', match: /^todo\b/, read: TASKS_READ, write: TASKS_WRITE },
  { area: 'OneNote', match: /^onenote\b/, read: NOTES_READ, write: NOTES_CREATE },
  { area: 'chats', match: /^chats\b/, read: ['Chat.ReadBasic', ...CHAT_READ], write: ['ChatMessage.Send', 'Chat.ReadWrite'] },
  { area: 'Teams', match: /^(teams|joinedteams)\b/, read: ['Team.ReadBasic.All', 'TeamSettings.Read.All', 'Channel.ReadBasic.All', 'ChannelMessage.Read.All', 'ChannelSettings.Read.All'], write: ['ChannelMessage.Send', 'TeamSettings.ReadWrite.All', 'ChannelSettings.ReadWrite.All', 'Channel.Create'] },
];
const READ_POSTS = /\/(getschedule|findmeetingtimes|search|query|delta)$/;

// Returns an error message if the connection lacks the scopes for a raw Graph call, else null.
// Paths outside the known areas are left to Microsoft, which checks the user's consent.
export function graphCallRefusal(method, path, granted) {
  const p = new URL(path, `${GRAPH_ORIGIN}/v1.0/`).pathname.toLowerCase().replace(/^\/(v1\.0|beta)/, '');
  const rest = p.replace(/^\/(me|users\/[^/]+)(?=\/|$)/, '').replace(/^\//, '');
  const area = AREAS.find((a) => a.match.test(rest));
  if (!area) return null;
  const reads = method === 'GET' || READ_POSTS.test(p);
  const needed = reads ? readOnly(area.read) : writable(area.write);
  if (hasAny(granted, needed)) return null;
  return `This connection has no ${reads ? 'read' : 'write'} access to ${area.area}. Call request_scopes with one of: ${needed.join(', ')}.`;
}
