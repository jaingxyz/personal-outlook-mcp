// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { z } from "zod";
import { graph } from "../graph.js";
import { config } from "../config.js";

// Graph IDs are opaque base64url-ish tokens. Some legitimately end in `=`;
// others don't. Normalizers between the model/client and our tool sometimes
// add or strip trailing `=` and Graph then rejects with "Id is malformed."
// We can't fix the input automatically (we don't know the canonical form),
// but we CAN detect that error and surface a useful message instead of the
// 14-character original.
function annotateGraphIdError(
  err: unknown,
  fields: Record<string, string>,
): never {
  const e = err as { code?: string; statusCode?: number; message?: string };
  const msg = e?.message ?? "";
  if (/id is malformed|invalid id|cast to.*failed/i.test(msg)) {
    const lines: string[] = [];
    for (const [name, value] of Object.entries(fields)) {
      const trailingEqs = value.length - value.replace(/=+$/, "").length;
      lines.push(
        `  ${name}: ends-with-=:${value.endsWith("=")}, trailing-eqs:${trailingEqs}, last-24:...${value.slice(-24)}`,
      );
    }
    throw new Error(
      `Graph rejected an id as malformed.\n${lines.join("\n")}\nIf the id was copied from a list/search call, try sending it byte-for-byte. Some normalizers add or strip base64 '=' padding that Graph rejects. Original error: ${msg}`,
    );
  }
  throw err;
}

// Wrap a Graph SDK promise so an "Id is malformed" rejection carries
// diagnostic context about which field(s) were used in the URL.
async function withIdContext<T>(
  fields: Record<string, string>,
  fn: () => Promise<T>,
): Promise<T> {
  try {
    return await fn();
  } catch (err) {
    annotateGraphIdError(err, fields);
  }
}

const dateTimeWithTz = z.object({
  dateTime: z
    .string()
    .min(1)
    .describe(
      "Local date-time, no offset, e.g. '2026-05-20T15:00:00'. The timeZone field controls interpretation.",
    ),
  timeZone: z
    .string()
    .min(1)
    .describe("IANA timezone, e.g. 'America/Los_Angeles' or 'UTC'."),
});

const attendeeSchema = z.object({
  email: z.string().email(),
  name: z.string().optional(),
  type: z
    .enum(["required", "optional", "resource"])
    .default("required")
    .describe("Attendee role. Defaults to required."),
});

type AttendeeInput = z.infer<typeof attendeeSchema>;

function toGraphAttendees(
  attendees: AttendeeInput[],
): Record<string, unknown>[] {
  return attendees.map((a) => ({
    emailAddress: { address: a.email, name: a.name },
    type: a.type,
  }));
}

function preferTzHeader(): string {
  return `outlook.timezone="${config.defaultTimeZone}"`;
}

const eventSelect = [
  "id",
  "subject",
  "start",
  "end",
  "isAllDay",
  "location",
  "organizer",
  "attendees",
  "isCancelled",
  "showAs",
  "responseStatus",
  "webLink",
  "bodyPreview",
  "type",
  "seriesMasterId",
  "recurrence",
].join(",");

export function summarizeEvent(
  e: Record<string, unknown>,
): Record<string, unknown> {
  return {
    id: e.id,
    subject: e.subject,
    start: e.start,
    end: e.end,
    isAllDay: e.isAllDay,
    location: (e.location as Record<string, unknown> | undefined)?.displayName,
    organizer: formatEmail(
      (
        e.organizer as
          | { emailAddress?: { name?: string; address?: string } }
          | undefined
      )?.emailAddress,
    ),
    attendees: ((e.attendees as Array<Record<string, unknown>>) ?? []).map(
      (a) => ({
        email: formatEmail(
          a.emailAddress as { name?: string; address?: string } | undefined,
        ),
        type: a.type,
        responseStatus: (a.status as { response?: string } | undefined)
          ?.response,
      }),
    ),
    isCancelled: e.isCancelled,
    showAs: e.showAs,
    myResponse: (e.responseStatus as { response?: string } | undefined)
      ?.response,
    type: e.type,
    seriesMasterId: e.seriesMasterId,
    webLink: e.webLink,
    preview: e.bodyPreview,
  };
}

export function formatEmail(
  ea: { name?: string; address?: string } | undefined | null,
): string | null {
  if (!ea) return null;
  if (ea.name && ea.address) return `${ea.name} <${ea.address}>`;
  return ea.address ?? ea.name ?? null;
}

// ---------- list_calendars ----------

export const listCalendarsSchema = z.object({});
export type ListCalendarsInput = z.infer<typeof listCalendarsSchema>;

export async function listCalendars(
  _input: ListCalendarsInput,
): Promise<unknown> {
  const res = await graph
    .api("/me/calendars")
    .select("id,name,isDefaultCalendar,canEdit,canShare,owner")
    .get();

  return {
    calendars: (res.value ?? []).map((c: Record<string, unknown>) => ({
      id: c.id,
      name: c.name,
      isDefault: c.isDefaultCalendar,
      canEdit: c.canEdit,
      owner: formatEmail(
        c.owner as { name?: string; address?: string } | undefined,
      ),
    })),
  };
}

// ---------- list_events (calendarView) ----------

export const listEventsSchema = z.object({
  start: z
    .string()
    .min(1)
    .describe(
      "Inclusive window start as ISO datetime, e.g. '2026-05-20T00:00:00' (treated as in defaultTimeZone if no offset).",
    ),
  end: z
    .string()
    .min(1)
    .describe("Exclusive window end. Same format as start."),
  calendarId: z
    .string()
    .optional()
    .describe("Specific calendar id. Defaults to the primary calendar."),
  limit: z.number().int().min(1).max(200).default(50),
});

export type ListEventsInput = z.infer<typeof listEventsSchema>;

export async function listEvents(input: ListEventsInput): Promise<unknown> {
  const path = input.calendarId
    ? `/me/calendars/${encodeURIComponent(input.calendarId)}/calendarView`
    : "/me/calendarView";

  const res = await withIdContext(
    input.calendarId ? { calendarId: input.calendarId } : {},
    () =>
      graph
        .api(path)
        .query({ startDateTime: input.start, endDateTime: input.end })
        .top(input.limit)
        .header("Prefer", preferTzHeader())
        .select(eventSelect)
        .orderby("start/dateTime")
        .get(),
  );

  return {
    timeZone: config.defaultTimeZone,
    events: (res.value ?? []).map(summarizeEvent),
  };
}

// ---------- read_event ----------

export const readEventSchema = z.object({
  eventId: z.string().min(1),
});
export type ReadEventInput = z.infer<typeof readEventSchema>;

export async function readEvent(input: ReadEventInput): Promise<unknown> {
  const e = await withIdContext({ eventId: input.eventId }, () =>
    graph
      .api(`/me/events/${encodeURIComponent(input.eventId)}`)
      .header("Prefer", preferTzHeader())
      .get(),
  );

  return {
    ...summarizeEvent(e),
    body: {
      contentType: (e.body as { contentType?: string } | undefined)
        ?.contentType,
      content: (e.body as { content?: string } | undefined)?.content,
    },
    recurrence: e.recurrence ?? null,
  };
}

// ---------- create_event ----------

export const createEventSchema = z.object({
  subject: z.string().min(1),
  start: dateTimeWithTz,
  end: dateTimeWithTz,
  attendees: z.array(attendeeSchema).default([]),
  location: z.string().optional(),
  body: z.string().optional(),
  bodyFormat: z.enum(["text", "html"]).default("text"),
  isOnlineMeeting: z
    .boolean()
    .default(false)
    .describe("If true, Graph attaches a Teams join link."),
  calendarId: z
    .string()
    .optional()
    .describe("Target calendar id. Defaults to primary."),
});

export type CreateEventInput = z.infer<typeof createEventSchema>;

export async function createEvent(input: CreateEventInput): Promise<unknown> {
  const path = input.calendarId
    ? `/me/calendars/${encodeURIComponent(input.calendarId)}/events`
    : "/me/events";

  const payload: Record<string, unknown> = {
    subject: input.subject,
    start: input.start,
    end: input.end,
    attendees: toGraphAttendees(input.attendees),
    isOnlineMeeting: input.isOnlineMeeting,
  };
  if (input.location) {
    payload.location = { displayName: input.location };
  }
  if (input.body) {
    payload.body = { contentType: input.bodyFormat, content: input.body };
  }

  const created = await withIdContext(
    input.calendarId ? { calendarId: input.calendarId } : {},
    () => graph.api(path).header("Prefer", preferTzHeader()).post(payload),
  );
  return summarizeEvent(created);
}

// ---------- update_event ----------

export const updateEventSchema = z.object({
  eventId: z.string().min(1),
  subject: z.string().optional(),
  start: dateTimeWithTz.optional(),
  end: dateTimeWithTz.optional(),
  attendees: z.array(attendeeSchema).optional(),
  location: z.string().optional(),
  body: z.string().optional(),
  bodyFormat: z.enum(["text", "html"]).default("text"),
});

export type UpdateEventInput = z.infer<typeof updateEventSchema>;

export async function updateEvent(input: UpdateEventInput): Promise<unknown> {
  // Single-occurrence edits of recurring series have a different code path
  // in Graph and additional gotchas. Refuse them in MVP.
  const existing = await withIdContext({ eventId: input.eventId }, () =>
    graph
      .api(`/me/events/${encodeURIComponent(input.eventId)}`)
      .select("type,seriesMasterId")
      .get(),
  );

  if (existing.type === "occurrence" || existing.type === "exception") {
    throw new Error(
      `Refusing to update a single occurrence of a recurring series (eventId=${input.eventId}, type=${existing.type}). Edit the series master instead, or cancel and recreate.`,
    );
  }

  const patch: Record<string, unknown> = {};
  if (input.subject !== undefined) patch.subject = input.subject;
  if (input.start !== undefined) patch.start = input.start;
  if (input.end !== undefined) patch.end = input.end;
  if (input.attendees !== undefined) {
    patch.attendees = toGraphAttendees(input.attendees);
  }
  if (input.location !== undefined) {
    patch.location = { displayName: input.location };
  }
  if (input.body !== undefined) {
    patch.body = { contentType: input.bodyFormat, content: input.body };
  }

  if (Object.keys(patch).length === 0) {
    return { ok: true, eventId: input.eventId, changed: [] };
  }

  await withIdContext({ eventId: input.eventId }, () =>
    graph.api(`/me/events/${encodeURIComponent(input.eventId)}`).patch(patch),
  );

  // Graph ignores Prefer: outlook.timezone on PATCH responses for events on
  // personal MSAs, so re-GET to render times in the configured zone.
  const updated = await withIdContext({ eventId: input.eventId }, () =>
    graph
      .api(`/me/events/${encodeURIComponent(input.eventId)}`)
      .header("Prefer", preferTzHeader())
      .select(eventSelect)
      .get(),
  );

  return {
    ok: true,
    eventId: input.eventId,
    changed: Object.keys(patch),
    event: summarizeEvent(updated),
  };
}

// ---------- cancel_event ----------

export const cancelEventSchema = z.object({
  eventId: z.string().min(1),
  comment: z
    .string()
    .optional()
    .describe(
      "Optional message included in the cancellation notice sent to attendees.",
    ),
  hardDelete: z
    .boolean()
    .default(false)
    .describe(
      "false (default): /cancel — sends notice to attendees and removes event for organizer. true: DELETE — removes event without sending notices. Use false if there are attendees.",
    ),
});

export type CancelEventInput = z.infer<typeof cancelEventSchema>;

export async function cancelEvent(input: CancelEventInput): Promise<unknown> {
  if (input.hardDelete) {
    await withIdContext({ eventId: input.eventId }, () =>
      graph.api(`/me/events/${encodeURIComponent(input.eventId)}`).delete(),
    );
    return { ok: true, eventId: input.eventId, mode: "hardDelete" };
  }

  await withIdContext({ eventId: input.eventId }, () =>
    graph
      .api(`/me/events/${encodeURIComponent(input.eventId)}/cancel`)
      .post(input.comment ? { Comment: input.comment } : {}),
  );

  return { ok: true, eventId: input.eventId, mode: "cancel" };
}

// ---------- respond_to_invite ----------

export const respondSchema = z.object({
  eventId: z.string().min(1),
  response: z.enum(["accept", "tentativelyAccept", "decline"]),
  comment: z.string().optional(),
  sendResponse: z
    .boolean()
    .default(true)
    .describe("Whether to email a response to the organizer."),
});

export type RespondInput = z.infer<typeof respondSchema>;

export async function respondToInvite(input: RespondInput): Promise<unknown> {
  const payload: Record<string, unknown> = {
    sendResponse: input.sendResponse,
  };
  if (input.comment) payload.comment = input.comment;

  await withIdContext({ eventId: input.eventId }, () =>
    graph
      .api(`/me/events/${encodeURIComponent(input.eventId)}/${input.response}`)
      .post(payload),
  );

  return { ok: true, eventId: input.eventId, response: input.response };
}
