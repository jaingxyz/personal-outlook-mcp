// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { describe, it, expect } from "vitest";
import { homedir } from "node:os";
import { formatRecipient, htmlToText } from "../src/tools/read.js";
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

describe("htmlToText", () => {
  it("drops script and style blocks entirely", () => {
    const out = htmlToText(
      "<style>.a{color:red}</style>Hi<script>alert(1)</script> there",
    );
    expect(out).not.toContain("color");
    expect(out).not.toContain("alert");
    expect(out).toContain("Hi");
    expect(out).toContain("there");
  });

  it("converts <br> and block tags to newlines", () => {
    expect(htmlToText("a<br>b")).toBe("a\nb");
    // Both the open and close of a block boundary break the line; adjacent
    // paragraphs end up separated by a blank line.
    expect(htmlToText("<p>one</p><p>two</p>")).toBe("one\n\ntwo");
  });

  it("decodes common entities and collapses excess blank lines", () => {
    expect(htmlToText("Tom &amp; Jerry &lt;3")).toBe("Tom & Jerry <3");
    expect(htmlToText("<p>a</p><p></p><p></p><p>b</p>")).toBe("a\n\nb");
  });

  it("strips HTML comments without leaking their tail (even when they contain '>')", () => {
    expect(htmlToText("<!-- promo: SAVE>20% -->X")).toBe("X");
    expect(htmlToText("a<!--[if mso]>junk<![endif]-->b")).toBe("ab");
  });

  it("separates adjacent blocks even when opening tags are unclosed", () => {
    // Nested divs without a symmetric close per line must not glue text.
    expect(htmlToText("<div>Line one<div>Line two</div>")).toBe(
      "Line one\nLine two",
    );
  });

  it("keeps a stray '<' that never closes as literal text", () => {
    expect(htmlToText("a < b")).toBe("a < b");
  });

  it("runs in linear time on adversarial input (no catastrophic backtracking)", () => {
    // Pathological body that pinned the old regex-chain version for seconds.
    const evil = "a < b ".repeat(100000); // 600KB of unterminated '<'
    const start = performance.now();
    htmlToText(evil);
    expect(performance.now() - start).toBeLessThan(1000);
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
      start: {
        dateTime: "2026-05-20T10:00:00",
        timeZone: "America/Los_Angeles",
      },
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
