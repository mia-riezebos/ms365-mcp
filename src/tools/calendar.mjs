import { userPath } from '../graph.mjs';
import { z, tool, enc, page, top, timeZone, dateTime, addr, bodyInput, bodyType, itemBody, compact } from './common.mjs';

const calendarOwner = z
  .string()
  .default('me')
  .describe('Whose calendar: "me", or the email address of a shared mailbox or of someone who shared their calendar with the signed-in user');
const tzHeader = (tz) => ({ Prefer: `outlook.timezone="${tz}"` });
const EVENT_FIELDS = 'id,subject,start,end,isAllDay,location,organizer,attendees,showAs,isCancelled,isOnlineMeeting,onlineMeeting,responseStatus,categories,bodyPreview,webLink,seriesMasterId,type';

const summarize = (e) => ({
  id: e.id,
  subject: e.subject,
  start: e.start?.dateTime,
  end: e.end?.dateTime,
  timeZone: e.start?.timeZone,
  isAllDay: e.isAllDay || undefined,
  location: e.location?.displayName || undefined,
  organizer: addr(e.organizer),
  attendees: e.attendees?.length ? e.attendees.map((a) => `${addr(a)} (${a.type}, ${a.status?.response})`) : undefined,
  showAs: e.showAs,
  isCancelled: e.isCancelled || undefined,
  onlineMeetingUrl: e.onlineMeeting?.joinUrl,
  myResponse: e.responseStatus?.response,
  categories: e.categories?.length ? e.categories : undefined,
  preview: e.bodyPreview || undefined,
  type: e.type !== 'singleInstance' ? e.type : undefined,
  webLink: e.webLink,
});

const attendeeList = (required = [], optional = []) => [
  ...required.map((address) => ({ emailAddress: { address }, type: 'required' })),
  ...optional.map((address) => ({ emailAddress: { address }, type: 'optional' })),
];

const eventFields = {
  subject: z.string().optional(),
  start: dateTime('Start').optional(),
  end: dateTime('End').optional(),
  time_zone: timeZone,
  is_all_day: z.boolean().optional().describe('All-day event; start and end must then be midnights'),
  location: z.string().optional(),
  body: bodyInput,
  body_type: bodyType,
  required_attendees: z.array(z.string().email()).optional(),
  optional_attendees: z.array(z.string().email()).optional(),
  is_online_meeting: z.boolean().optional().describe('Add a Teams meeting link'),
  show_as: z.enum(['free', 'tentative', 'busy', 'oof', 'workingElsewhere']).optional(),
  categories: z.array(z.string()).optional(),
  reminder_minutes: z.number().int().min(0).optional(),
};

const eventBody = (a) =>
  compact({
    subject: a.subject,
    start: a.start && { dateTime: a.start, timeZone: a.time_zone },
    end: a.end && { dateTime: a.end, timeZone: a.time_zone },
    isAllDay: a.is_all_day,
    location: a.location !== undefined ? { displayName: a.location } : undefined,
    body: itemBody(a.body, a.body_type),
    attendees: a.required_attendees || a.optional_attendees ? attendeeList(a.required_attendees, a.optional_attendees) : undefined,
    isOnlineMeeting: a.is_online_meeting,
    onlineMeetingProvider: a.is_online_meeting ? 'teamsForBusiness' : undefined,
    showAs: a.show_as,
    categories: a.categories,
    reminderMinutesBeforeStart: a.reminder_minutes,
    isReminderOn: a.reminder_minutes !== undefined ? true : undefined,
  });

export function registerCalendar(server, graph) {
  tool(server, 'list_calendars', { title: 'List calendars', description: 'Lists the calendars of a user or shared mailbox.', input: { owner: calendarOwner } }, async ({ owner }) =>
    page(await graph('GET', `${userPath(owner)}/calendars`, { query: { $select: 'id,name,color,isDefaultCalendar,canEdit,owner' } })),
  );

  tool(
    server,
    'list_events',
    {
      title: 'List calendar events',
      description: 'Lists events between two times (recurring events expanded into occurrences), soonest first.',
      input: {
        owner: calendarOwner,
        calendar_id: z.string().optional().describe('Calendar id; omit for the default calendar'),
        start: dateTime('Window start'),
        end: dateTime('Window end'),
        time_zone: timeZone,
        filter: z.string().optional().describe('OData $filter, e.g. "contains(subject,\'review\')"'),
        top: top(100, 50),
      },
    },
    async ({ owner, calendar_id, start, end, time_zone, filter, top }) => {
      const base = `${userPath(owner)}${calendar_id ? `/calendars/${enc(calendar_id)}` : '/calendar'}/calendarView`;
      const res = await graph('GET', base, {
        query: { startDateTime: start, endDateTime: end, $top: top, $orderby: 'start/dateTime', $select: EVENT_FIELDS, $filter: filter },
        headers: tzHeader(time_zone),
      });
      return page(res, summarize);
    },
  );

  tool(
    server,
    'get_event',
    { title: 'Read an event', description: 'Gets one event with its full body.', input: { owner: calendarOwner, event_id: z.string(), time_zone: timeZone } },
    async ({ owner, event_id, time_zone }) => {
      const e = await graph('GET', `${userPath(owner)}/events/${enc(event_id)}`, { headers: { ...tzHeader(time_zone), Prefer: `outlook.timezone="${time_zone}", outlook.body-content-type="text"` } });
      return { ...summarize(e), body: e.body?.content, preview: undefined, recurrence: e.recurrence || undefined };
    },
  );

  tool(
    server,
    'create_event',
    {
      title: 'Create an event',
      description: 'Creates a calendar event. With attendees, Outlook sends them invitations.',
      kind: 'destructive',
      input: { owner: calendarOwner, calendar_id: z.string().optional(), ...eventFields, subject: z.string(), start: dateTime('Start'), end: dateTime('End') },
    },
    async (a) => {
      const path = `${userPath(a.owner)}${a.calendar_id ? `/calendars/${enc(a.calendar_id)}` : '/calendar'}/events`;
      return summarize(await graph('POST', path, { body: eventBody(a), headers: tzHeader(a.time_zone) }));
    },
  );

  tool(
    server,
    'update_event',
    {
      title: 'Update an event',
      description: 'Changes fields of an event; only the fields given change. Attendees get an update if the organizer changes a meeting. Attendee lists replace the current ones.',
      kind: 'destructive',
      input: { owner: calendarOwner, event_id: z.string(), ...eventFields },
    },
    async (a) => summarize(await graph('PATCH', `${userPath(a.owner)}/events/${enc(a.event_id)}`, { body: eventBody(a), headers: tzHeader(a.time_zone) })),
  );

  tool(
    server,
    'delete_event',
    {
      title: 'Cancel or delete an event',
      description: 'As organizer of a meeting with attendees, cancels it and notifies them (with an optional comment); otherwise deletes the event from the calendar.',
      kind: 'destructive',
      input: { owner: calendarOwner, event_id: z.string(), comment: z.string().optional() },
    },
    async ({ owner, event_id, comment }) => {
      const path = `${userPath(owner)}/events/${enc(event_id)}`;
      const e = await graph('GET', path, { query: { $select: 'isOrganizer,attendees' } });
      if (e.isOrganizer && e.attendees?.length) {
        await graph('POST', `${path}/cancel`, { body: { comment: comment || '' } });
        return { cancelled: true, notified: e.attendees.length };
      }
      await graph('DELETE', path);
      return { deleted: true };
    },
  );

  tool(
    server,
    'respond_to_event',
    {
      title: 'Respond to an invitation',
      description: 'Accepts, tentatively accepts or declines a meeting invitation.',
      kind: 'destructive',
      input: {
        owner: calendarOwner,
        event_id: z.string(),
        response: z.enum(['accept', 'tentativelyAccept', 'decline']),
        comment: z.string().optional(),
        send_response: z.boolean().default(true).describe('Tell the organizer'),
      },
    },
    async ({ owner, event_id, response, comment, send_response }) => {
      await graph('POST', `${userPath(owner)}/events/${enc(event_id)}/${response}`, { body: compact({ comment, sendResponse: send_response }) });
      return { responded: response };
    },
  );

  tool(
    server,
    'get_free_busy',
    {
      title: 'Check free/busy',
      description: 'Gets free/busy blocks for people, rooms or shared mailboxes between two times.',
      input: {
        addresses: z.array(z.string().email()).min(1).max(20),
        start: dateTime('Window start'),
        end: dateTime('Window end'),
        time_zone: timeZone,
        interval_minutes: z.number().int().min(5).max(1440).default(30),
      },
    },
    async ({ addresses, start, end, time_zone, interval_minutes }) => {
      const res = await graph('POST', '/me/calendar/getSchedule', {
        body: { schedules: addresses, startTime: { dateTime: start, timeZone: time_zone }, endTime: { dateTime: end, timeZone: time_zone }, availabilityViewInterval: interval_minutes },
        headers: tzHeader(time_zone),
      });
      return res.value.map((s) => ({
        address: s.scheduleId,
        error: s.error?.message,
        availabilityView: s.availabilityView,
        busy: s.scheduleItems?.map((i) => ({ status: i.status, start: i.start?.dateTime, end: i.end?.dateTime, subject: i.subject || undefined })),
        workingHours: s.workingHours && { days: s.workingHours.daysOfWeek, start: s.workingHours.startTime, end: s.workingHours.endTime, timeZone: s.workingHours.timeZone?.name },
      }));
    },
  );
}
