// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { describe, it, expect, beforeEach, vi } from "vitest";
import { makeMockGraph } from "./_mockGraph.js";

const mock = makeMockGraph();

vi.mock("../src/graph.js", () => ({
  graph: { api: mock.api },
  getMe: vi.fn(),
}));

beforeEach(() => {
  mock.calls.length = 0;
  mock.responses.length = 0;
  mock.api.mockClear();
});

describe("calendar tools", () => {
  it("listEvents uses /me/calendarView with start/end query and Prefer header", async () => {
    const { listEvents } = await import("../src/tools/calendar.js");
    mock.responses.push({ method: "get", value: { value: [] } });
    await listEvents({
      start: "2026-05-20T00:00:00",
      end: "2026-05-21T00:00:00",
      limit: 10,
    });
    const c = mock.calls[0];
    expect(c.path).toBe("/me/calendarView");
    expect(c.query).toEqual({
      startDateTime: "2026-05-20T00:00:00",
      endDateTime: "2026-05-21T00:00:00",
    });
    expect(c.headers["Prefer"]).toMatch(/^outlook\.timezone=/);
    expect(c.orderby).toBe("start/dateTime");
  });

  it("listEvents scopes to a calendar id when given", async () => {
    const { listEvents } = await import("../src/tools/calendar.js");
    mock.responses.push({ method: "get", value: { value: [] } });
    await listEvents({
      start: "x",
      end: "y",
      calendarId: "cal1",
      limit: 5,
    });
    expect(mock.calls[0].path).toBe("/me/calendars/cal1/calendarView");
  });

  it("createEvent POSTs with Prefer header so response is in user TZ", async () => {
    const { createEvent } = await import("../src/tools/calendar.js");
    mock.responses.push({
      method: "post",
      value: {
        id: "e1",
        subject: "test",
        start: { dateTime: "x", timeZone: "America/Los_Angeles" },
        end: { dateTime: "x", timeZone: "America/Los_Angeles" },
      },
    });
    await createEvent({
      subject: "test",
      start: {
        dateTime: "2026-05-20T10:00:00",
        timeZone: "America/Los_Angeles",
      },
      end: { dateTime: "2026-05-20T10:15:00", timeZone: "America/Los_Angeles" },
      attendees: [],
      bodyFormat: "text",
      isOnlineMeeting: false,
    });
    expect(mock.calls[0].headers["Prefer"]).toMatch(/^outlook\.timezone=/);
    const body = mock.calls[0].body as Record<string, unknown>;
    expect(body.subject).toBe("test");
    expect(body.attendees).toEqual([]);
  });

  it("updateEvent PATCHes then re-GETs with Prefer header (works around Graph PATCH-Prefer bug)", async () => {
    const { updateEvent } = await import("../src/tools/calendar.js");
    // existing-event check
    mock.responses.push({
      method: "get",
      value: { type: "singleInstance", seriesMasterId: null },
    });
    // PATCH
    mock.responses.push({ method: "patch", value: {} });
    // Re-GET in TZ
    mock.responses.push({
      method: "get",
      value: {
        id: "e1",
        subject: "renamed",
        start: { dateTime: "x", timeZone: "America/Los_Angeles" },
        end: { dateTime: "x", timeZone: "America/Los_Angeles" },
      },
    });
    const out = (await updateEvent({
      eventId: "e1",
      subject: "renamed",
      bodyFormat: "text",
    })) as Record<string, unknown>;

    expect(mock.calls.map((c) => c.method)).toEqual(["get", "patch", "get"]);
    // The re-GET must have the Prefer header so display TZ is user's, not UTC.
    expect(mock.calls[2].headers["Prefer"]).toMatch(/^outlook\.timezone=/);
    expect(out.changed).toEqual(["subject"]);
  });

  it("updateEvent refuses single occurrences of recurring series", async () => {
    const { updateEvent } = await import("../src/tools/calendar.js");
    mock.responses.push({
      method: "get",
      value: { type: "occurrence", seriesMasterId: "series1" },
    });
    await expect(
      updateEvent({
        eventId: "occ1",
        subject: "x",
        bodyFormat: "text",
      }),
    ).rejects.toThrow(/single occurrence/i);
  });

  it("updateEvent is a no-op when nothing is set", async () => {
    const { updateEvent } = await import("../src/tools/calendar.js");
    mock.responses.push({
      method: "get",
      value: { type: "singleInstance" },
    });
    const out = (await updateEvent({
      eventId: "e1",
      bodyFormat: "text",
    })) as Record<string, unknown>;
    expect(out.changed).toEqual([]);
    // Should not have queued a PATCH; only the type-check GET.
    expect(mock.calls).toHaveLength(1);
  });

  it("cancelEvent in cancel mode POSTs /cancel with optional comment", async () => {
    const { cancelEvent } = await import("../src/tools/calendar.js");
    mock.responses.push({ method: "post", value: undefined });
    await cancelEvent({
      eventId: "e1",
      comment: "rescheduling",
      hardDelete: false,
    });
    expect(mock.calls[0].path).toBe("/me/events/e1/cancel");
    expect(mock.calls[0].body).toEqual({ Comment: "rescheduling" });
  });

  it("cancelEvent in hardDelete mode DELETEs", async () => {
    const { cancelEvent } = await import("../src/tools/calendar.js");
    mock.responses.push({ method: "delete", value: undefined });
    const out = (await cancelEvent({
      eventId: "e1",
      hardDelete: true,
    })) as Record<string, unknown>;
    expect(mock.calls[0].method).toBe("delete");
    expect(out.mode).toBe("hardDelete");
  });

  it("respondToInvite POSTs to /accept|/decline|/tentativelyAccept", async () => {
    const { respondToInvite } = await import("../src/tools/calendar.js");
    for (const r of ["accept", "decline", "tentativelyAccept"] as const) {
      mock.responses.push({ method: "post", value: undefined });
      await respondToInvite({
        eventId: "e1",
        response: r,
        sendResponse: true,
      });
    }
    expect(mock.calls.map((c) => c.path)).toEqual([
      "/me/events/e1/accept",
      "/me/events/e1/decline",
      "/me/events/e1/tentativelyAccept",
    ]);
  });

  it("createEvent annotates 'Id is malformed' with calendarId diagnostics (regression)", async () => {
    // Repro: a stray trailing '=' was appended to a calendarId by an
    // upstream normalizer, Graph rejected with the unhelpful 14-char string
    // "Id is malformed.". The wrapper should surface the suspect field +
    // its trailing-equals state so the model/user can see what to fix.
    const { createEvent } = await import("../src/tools/calendar.js");
    const graphErr = Object.assign(new Error("Id is malformed."), {
      statusCode: 400,
    });
    mock.responses.push({ method: "post", value: graphErr, throws: true });

    const badId = "AQMkADAwSomeBase64UrlIshThing=";
    await expect(
      createEvent({
        subject: "x",
        start: { dateTime: "2026-06-04T08:30:00", timeZone: "UTC" },
        end: { dateTime: "2026-06-04T09:00:00", timeZone: "UTC" },
        attendees: [],
        bodyFormat: "text",
        isOnlineMeeting: false,
        calendarId: badId,
      }),
    ).rejects.toThrow(
      /Graph rejected an id as malformed[\s\S]*calendarId.*ends-with-=:true.*trailing-eqs:1/,
    );
  });

  it("createEvent passes through non-malformed-id errors unchanged", async () => {
    const { createEvent } = await import("../src/tools/calendar.js");
    const otherErr = new Error("Throttled");
    mock.responses.push({ method: "post", value: otherErr, throws: true });

    await expect(
      createEvent({
        subject: "x",
        start: { dateTime: "2026-06-04T08:30:00", timeZone: "UTC" },
        end: { dateTime: "2026-06-04T09:00:00", timeZone: "UTC" },
        attendees: [],
        bodyFormat: "text",
        isOnlineMeeting: false,
        calendarId: "anything",
      }),
    ).rejects.toThrow("Throttled");
  });
});
