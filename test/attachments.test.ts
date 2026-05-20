// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import { mkdtempSync, rmSync, readFileSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { makeMockGraph } from "./_mockGraph.js";

const mock = makeMockGraph();

vi.mock("../src/graph.js", () => ({
  graph: { api: mock.api },
  getMe: vi.fn(),
}));

let tmp: string;

beforeEach(() => {
  mock.calls.length = 0;
  mock.responses.length = 0;
  mock.api.mockClear();
  tmp = mkdtempSync(join(tmpdir(), "mcp-att-"));
});

afterEach(() => {
  rmSync(tmp, { recursive: true, force: true });
});

describe("attachment tools", () => {
  it("listAttachments selects metadata fields", async () => {
    const { listAttachments } = await import("../src/tools/attachments.js");
    mock.responses.push({
      method: "get",
      value: {
        value: [
          {
            id: "a1",
            name: "report.pdf",
            contentType: "application/pdf",
            size: 1234,
            isInline: false,
            "@odata.type": "#microsoft.graph.fileAttachment",
          },
        ],
      },
    });
    const out = (await listAttachments({ messageId: "m1" })) as Record<
      string,
      unknown
    >;
    expect(mock.calls[0].path).toBe("/me/messages/m1/attachments");
    expect(mock.calls[0].select).toContain("contentType");
    const list = out.attachments as Array<Record<string, unknown>>;
    expect(list[0]).toMatchObject({
      id: "a1",
      name: "report.pdf",
      kind: "#microsoft.graph.fileAttachment",
    });
  });

  it("downloadAttachment writes the decoded file to disk and returns metadata", async () => {
    const { downloadAttachment } = await import(
      "../src/tools/attachments.js"
    );
    const content = Buffer.from("hello world");
    mock.responses.push({
      method: "get",
      value: {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: "greeting.txt",
        contentType: "text/plain",
        contentBytes: content.toString("base64"),
      },
    });

    const out = (await downloadAttachment({
      messageId: "m1",
      attachmentId: "a1",
      destDir: tmp,
    })) as Record<string, unknown>;

    expect(out.ok).toBe(true);
    expect(out.size).toBe(content.length);
    expect(out.path).toBe(join(tmp, "greeting.txt"));
    const onDisk = readFileSync(out.path as string);
    expect(onDisk.equals(content)).toBe(true);
    expect(statSync(out.path as string).size).toBe(content.length);
  });

  it("downloadAttachment rejects non-fileAttachment types", async () => {
    const { downloadAttachment } = await import(
      "../src/tools/attachments.js"
    );
    mock.responses.push({
      method: "get",
      value: {
        "@odata.type": "#microsoft.graph.itemAttachment",
        name: "weird.eml",
      },
    });
    await expect(
      downloadAttachment({
        messageId: "m1",
        attachmentId: "a1",
        destDir: tmp,
      }),
    ).rejects.toThrow(/Unsupported attachment type/);
  });

  it("downloadAttachment honors filename override and sanitizes it", async () => {
    const { downloadAttachment } = await import(
      "../src/tools/attachments.js"
    );
    mock.responses.push({
      method: "get",
      value: {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: "original.txt",
        contentType: "text/plain",
        contentBytes: Buffer.from("x").toString("base64"),
      },
    });
    const out = (await downloadAttachment({
      messageId: "m1",
      attachmentId: "a1",
      destDir: tmp,
      filename: "subdir/escape.txt",
    })) as Record<string, unknown>;
    // Slash is sanitized to _, so the file lands in tmp, not in tmp/subdir.
    expect(out.name).toBe("subdir_escape.txt");
    expect(out.path).toBe(join(tmp, "subdir_escape.txt"));
  });
});
