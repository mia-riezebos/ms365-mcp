import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
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
export function createMcpServer(graph) {
  const server = new McpServer({ name: 'ms365', version: '0.2.0' });
  for (const register of [registerCore, registerMail, registerCalendar, registerContacts, registerTeams, registerPlanner, registerTodo, registerFiles, registerOneNote]) register(server, graph);
  return server;
}
