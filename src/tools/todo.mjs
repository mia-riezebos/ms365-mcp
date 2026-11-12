import { z, tool, enc, page, top, compact, itemBody } from './common.mjs';

const summarize = (t) => ({
  id: t.id,
  title: t.title,
  status: t.status,
  importance: t.importance !== 'normal' ? t.importance : undefined,
  due: t.dueDateTime?.dateTime?.slice(0, 10),
  reminder: t.isReminderOn ? t.reminderDateTime?.dateTime : undefined,
  body: t.body?.content?.trim() || undefined,
  categories: t.categories?.length ? t.categories : undefined,
  created: t.createdDateTime,
  completed: t.completedDateTime?.dateTime,
});

const fields = {
  title: z.string().optional(),
  body: z.string().optional(),
  due_date: z.string().optional().describe('YYYY-MM-DD'),
  reminder: z.string().optional().describe('Reminder time, e.g. 2026-11-20T09:00:00 (in time_zone)'),
  time_zone: z.string().default('Europe/Amsterdam'),
  importance: z.enum(['low', 'normal', 'high']).optional(),
  status: z.enum(['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred']).optional(),
  categories: z.array(z.string()).optional(),
};
const taskBody = (a) =>
  compact({
    title: a.title,
    body: itemBody(a.body, 'text'),
    dueDateTime: a.due_date && { dateTime: `${a.due_date}T00:00:00`, timeZone: a.time_zone },
    reminderDateTime: a.reminder && { dateTime: a.reminder, timeZone: a.time_zone },
    isReminderOn: a.reminder ? true : undefined,
    importance: a.importance,
    status: a.status,
    categories: a.categories,
  });

export function registerTodo(server, graph) {
  tool(server, 'list_todo_lists', { title: 'List To Do lists', description: 'Lists Microsoft To Do task lists.' }, async () =>
    page(await graph('GET', '/me/todo/lists'), (l) => ({ id: l.id, name: l.displayName, wellknown: l.wellknownListName !== 'none' ? l.wellknownListName : undefined, shared: l.isShared || undefined })),
  );

  tool(
    server,
    'list_todo_tasks',
    { title: 'List To Do tasks', description: 'Lists tasks in a To Do list; open tasks only unless include_completed.', input: { list_id: z.string(), include_completed: z.boolean().default(false), top: top(100, 50) } },
    async ({ list_id, include_completed, top }) =>
      page(await graph('GET', `/me/todo/lists/${enc(list_id)}/tasks`, { query: { $top: top, $filter: include_completed ? undefined : "status ne 'completed'" } }), summarize),
  );

  tool(server, 'create_todo_task', { title: 'Create a To Do task', description: 'Adds a task to a To Do list.', kind: 'write', input: { list_id: z.string(), ...fields, title: z.string() } }, async (a) =>
    summarize(await graph('POST', `/me/todo/lists/${enc(a.list_id)}/tasks`, { body: taskBody(a) })),
  );

  tool(
    server,
    'update_todo_task',
    { title: 'Update a To Do task', description: 'Changes the given fields of a To Do task (status completed to tick it off).', kind: 'write', input: { list_id: z.string(), task_id: z.string(), ...fields } },
    async (a) => summarize(await graph('PATCH', `/me/todo/lists/${enc(a.list_id)}/tasks/${enc(a.task_id)}`, { body: taskBody(a) })),
  );

  tool(server, 'delete_todo_task', { title: 'Delete a To Do task', description: 'Deletes a To Do task.', kind: 'destructive', input: { list_id: z.string(), task_id: z.string() } }, async ({ list_id, task_id }) => {
    await graph('DELETE', `/me/todo/lists/${enc(list_id)}/tasks/${enc(task_id)}`);
    return { deleted: true };
  });
}
