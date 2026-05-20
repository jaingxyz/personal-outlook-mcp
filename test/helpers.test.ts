// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { describe, it, expect } from "vitest";
import { homedir } from "node:os";
import { formatRecipient } from "../src/tools/read.js";
import { sanitizeFilename, expandHome } from "../src/tools/attachments.js";
import { formatEmail, summarizeEvent } from "../src/tools/calendar.js";

describe("formatRecipient", () => {
  it("returns 'Name <addr>' when both fields present", () => {
    expect(
      formatRecipient({ emailAddress: { name: "Alice", address: "a@b.com" } }),
    ).toBe("Alice <a@b.com>");
  });

  it("falls back to address when name missing", () => {
    expect(formatRecipient({ emailAddress: { address: "a@b.com" } })).toBe(
      "a@b.com",
    );
  });

  it("falls back to name when address missing", () => {
    expect(formatRecipient({ emailAddress: { name: "Alice" } })).toBe("Alice");
  });

  it("returns null for nullish or empty input", () => {
    expect(formatRecipient(null)).toBeNull();
    expect(formatRecipient(undefined)).toBeNull();
    expect(formatRecipient({})).toBeNull();
    expect(formatRecipient({ emailAddress: {} })).toBeNull();
  });
});

describe("sanitizeFilename", () => {
  it("strips path separators (defense-in-depth: result is always a single filename, never traverses dirs when joined)", () => {
    // Forward and backslash become _; .. survives but is harmless once path.resolve treats result as a filename.
    expect(sanitizeFilename("../../etc/passwd")).toBe(".._.._etc_passwd");
    expect(sanitizeFilename("dir\\file.txt")).toBe("dir_file.txt");
    // No / remains, so the result cannot escape the chosen destDir.
    expect(sanitizeFilename("../../etc/passwd")).not.toContain("/");
    expect(sanitizeFilename("dir\\file.txt")).not.toContain("\\");
  });

  it("strips control characters", () => {
    expect(sanitizeFilename("a\x00b\x1fc")).toBe("a_b_c");
  });

  it("preserves spaces and normal punctuation", () => {
    expect(sanitizeFilename("Trip itinerary - May 20, 2026.pdf")).toBe(
      "Trip itinerary - May 20, 2026.pdf",
    );
  });

  it("returns 'attachment' for empty or whitespace-only names", () => {
    expect(sanitizeFilename("")).toBe("attachment");
    expect(sanitizeFilename("   ")).toBe("attachment");
  });
});

describe("expandHome", () => {
  it("expands bare ~", () => {
    expect(expandHome("~")).toBe(homedir());
  });

  it("expands ~/ prefix", () => {
    expect(expandHome("~/Downloads")).toBe(`${homedir()}/Downloads`);
  });

  it("leaves absolute paths untouched", () => {
    expect(expandHome("/tmp/foo")).toBe("/tmp/foo");
  });

  it("does not expand mid-path tildes", () => {
    expect(expandHome("/etc/~weird")).toBe("/etc/~weird");
  });
});

describe("formatEmail", () => {
  it("handles both fields", () => {
    expect(formatEmail({ name: "Bob", address: "b@c.com" })).toBe(
      "Bob <b@c.com>",
    );
  });

  it("returns null for nullish", () => {
    expect(formatEmail(undefined)).toBeNull();
    expect(formatEmail(null)).toBeNull();
  });
});

describe("summarizeEvent", () => {
  it("flattens organizer and attendees into displayable strings", () => {
    const raw = {
      id: "evt1",
      subject: "Standup",
      start: { dateTime: "2026-05-20T10:00:00", timeZone: "America/Los_Angeles" },
      end: { dateTime: "2026-05-20T10:15:00", timeZone: "America/Los_Angeles" },
      isAllDay: false,
      location: { displayName: "Room 1" },
      organizer: {
        emailAddress: { name: "Alice", address: "alice@example.com" },
      },
      attendees: [
        {
          emailAddress: { name: "Bob", address: "bob@example.com" },
          type: "required",
          status: { response: "accepted" },
        },
      ],
      isCancelled: false,
      showAs: "busy",
      responseStatus: { response: "accepted" },
      type: "singleInstance",
      webLink: "https://example.com",
      bodyPreview: "Daily standup",
    };

    const out = summarizeEvent(raw) as Record<string, unknown>;
    expect(out.id).toBe("evt1");
    expect(out.location).toBe("Room 1");
    expect(out.organizer).toBe("Alice <alice@example.com>");
    expect(out.myResponse).toBe("accepted");
    expect((out.attendees as Array<Record<string, unknown>>)[0]).toEqual({
      email: "Bob <bob@example.com>",
      type: "required",
      responseStatus: "accepted",
    });
  });

  it("handles missing optional fields gracefully", () => {
    const out = summarizeEvent({
      id: "x",
      subject: "no frills",
    }) as Record<string, unknown>;
    expect(out.location).toBeUndefined();
    expect(out.organizer).toBeNull();
    expect(out.attendees).toEqual([]);
    expect(out.myResponse).toBeUndefined();
  });
});
