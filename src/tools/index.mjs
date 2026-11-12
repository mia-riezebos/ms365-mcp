import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { toolAllowed } from '../scopes.mjs';
import { registerCore } from './core.mjs';
import { registerMail } from './mail.mjs';
import { registerCalendar } from './calendar.mjs';
import { registerContacts } from './contacts.mjs';
import { registerTeams } from './teams.mjs';
import { registerPlanner } from './planner.mjs';
import { registerTodo } from './todo.mjs';
import { registerFiles } from './files.mjs';
import { registerOneNote } from './onenote.mjs';

// graph: (method, path, options) => Graph response, as returned by graphClient().
// conn: the connection's granted scopes and helpers (see registerCore); only tools its scopes allow are offered.
export function createMcpServer(graph, conn) {
  const server = new McpServer({ name: 'ms365', version: '0.3.0' });
  const register = server.registerTool.bind(server);
  server.registerTool = (name, ...rest) => (toolAllowed(name, conn.granted) ? register(name, ...rest) : undefined);
  for (const fn of [registerCore, registerMail, registerCalendar, registerContacts, registerTeams, registerPlanner, registerTodo, registerFiles, registerOneNote]) fn(server, graph, conn);
  return server;
}
