import { z, tool, enc, page, compact } from './common.mjs';

const summarizeTask = (t) => ({
  id: t.id,
  title: t.title,
  planId: t.planId,
  bucketId: t.bucketId,
  percentComplete: t.percentComplete,
  priority: t.priority,
  start: t.startDateTime || undefined,
  due: t.dueDateTime || undefined,
  assignees: Object.keys(t.assignments || {}),
  hasDescription: t.hasDescription,
  created: t.createdDateTime,
  completed: t.completedDateTime || undefined,
});

const taskFields = {
  title: z.string().optional(),
  bucket_id: z.string().optional(),
  due: z.string().optional().describe('Due date-time, ISO 8601 with offset, e.g. 2026-11-20T17:00:00+01:00'),
  start: z.string().optional(),
  percent_complete: z.number().int().min(0).max(100).optional().describe('0 not started, 50 in progress, 100 done'),
  priority: z.number().int().min(0).max(10).optional().describe('1 urgent, 3 important, 5 medium, 9 low'),
  assignee_ids: z.array(z.string()).optional().describe('Entra user ids to assign (adds to current assignees)'),
};
const taskBody = (a) =>
  compact({
    title: a.title,
    bucketId: a.bucket_id,
    dueDateTime: a.due,
    startDateTime: a.start,
    percentComplete: a.percent_complete,
    priority: a.priority,
    assignments: a.assignee_ids?.length
      ? Object.fromEntries(a.assignee_ids.map((id) => [id, { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' }]))
      : undefined,
  });

export function registerPlanner(server, graph) {
  tool(
    server,
    'list_plans',
    { title: 'List Planner plans', description: 'Lists Planner plans shared with the signed-in user, or the plans of one Microsoft 365 group.', input: { group_id: z.string().optional() } },
    async ({ group_id }) => page(await graph('GET', group_id ? `/groups/${enc(group_id)}/planner/plans` : '/me/planner/plans'), (p) => ({ id: p.id, title: p.title, groupId: p.container?.containerId || p.owner })),
  );

  tool(server, 'list_buckets', { title: 'List Planner buckets', description: 'Lists the buckets (columns) of a plan.', input: { plan_id: z.string() } }, async ({ plan_id }) =>
    page(await graph('GET', `/planner/plans/${enc(plan_id)}/buckets`), (b) => ({ id: b.id, name: b.name })),
  );

  tool(
    server,
    'list_planner_tasks',
    { title: 'List Planner tasks', description: 'Lists the tasks of a plan, or all Planner tasks assigned to the signed-in user.', input: { plan_id: z.string().optional(), include_completed: z.boolean().default(false) } },
    async ({ plan_id, include_completed }) => {
      const res = page(await graph('GET', plan_id ? `/planner/plans/${enc(plan_id)}/tasks` : '/me/planner/tasks'), summarizeTask);
      if (!include_completed) res.items = res.items.filter((t) => t.percentComplete < 100);
      res.count = res.items.length;
      return res;
    },
  );

  tool(server, 'get_planner_task', { title: 'Read a Planner task', description: 'Gets a Planner task with its description and checklist.', input: { task_id: z.string() } }, async ({ task_id }) => {
    const [t, d] = await Promise.all([graph('GET', `/planner/tasks/${enc(task_id)}`), graph('GET', `/planner/tasks/${enc(task_id)}/details`)]);
    return { ...summarizeTask(t), description: d.description, checklist: Object.values(d.checklist || {}).map((c) => ({ title: c.title, done: c.isChecked })) };
  });

  const setDescription = async (task_id, description) => {
    const d = await graph('GET', `/planner/tasks/${enc(task_id)}/details`);
    await graph('PATCH', `/planner/tasks/${enc(task_id)}/details`, { body: { description }, headers: { 'If-Match': d['@odata.etag'] } });
  };

  tool(
    server,
    'create_planner_task',
    { title: 'Create a Planner task', description: 'Adds a task to a plan.', kind: 'write', input: { plan_id: z.string(), ...taskFields, title: z.string(), description: z.string().optional() } },
    async (a) => {
      const t = await graph('POST', '/planner/tasks', { body: { planId: a.plan_id, ...taskBody(a) } });
      if (a.description) await setDescription(t.id, a.description);
      return summarizeTask(t);
    },
  );

  tool(
    server,
    'update_planner_task',
    { title: 'Update a Planner task', description: 'Changes the given fields of a Planner task (set percent_complete 100 to complete it).', kind: 'write', input: { task_id: z.string(), ...taskFields, description: z.string().optional() } },
    async (a) => {
      const body = taskBody(a);
      if (Object.keys(body).length) {
        const t = await graph('GET', `/planner/tasks/${enc(a.task_id)}`);
        await graph('PATCH', `/planner/tasks/${enc(a.task_id)}`, { body, headers: { 'If-Match': t['@odata.etag'] } });
      }
      if (a.description !== undefined) await setDescription(a.task_id, a.description);
      return summarizeTask(await graph('GET', `/planner/tasks/${enc(a.task_id)}`));
    },
  );

  tool(server, 'delete_planner_task', { title: 'Delete a Planner task', description: 'Deletes a Planner task.', kind: 'destructive', input: { task_id: z.string() } }, async ({ task_id }) => {
    const t = await graph('GET', `/planner/tasks/${enc(task_id)}`);
    await graph('DELETE', `/planner/tasks/${enc(task_id)}`, { headers: { 'If-Match': t['@odata.etag'] } });
    return { deleted: true };
  });
}
