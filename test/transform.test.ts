import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import {
  buildPageHtml,
  buildSummaryHtml,
  formatTimestamp,
  parseActionItems,
  parseInsightNotes,
  parseVtt,
} from "../src/lib/transform-core.ts";

const sampleVtt = readFileSync(fileURLToPath(new URL("fixtures/sample.vtt", import.meta.url)), "utf8");

describe("parseVtt", () => {
  it("parses cues and merges consecutive turns from the same speaker", () => {
    const turns = parseVtt(sampleVtt);
    expect(turns).toHaveLength(4);
    expect(turns[0]).toEqual({
      speaker: "Venkat Yenduri",
      startSeconds: 3,
      text: "Hello everyone, thanks for joining. Let's get started with the roadmap.",
    });
  });

  it("decodes HTML entities and keeps voice-class speakers", () => {
    const turns = parseVtt(sampleVtt);
    expect(turns[1]?.text).toBe("Sounds good. The Q3 numbers & targets look strong.");
    expect(turns[2]?.speaker).toBe("Bob O'Brien");
    expect(turns[2]?.text).toBe("I'll take the action item on pricing.");
  });

  it("handles mm:ss timestamps and missing closing tags", () => {
    const vtt = "WEBVTT\n\n00:05.000 --> 00:07.000\n<v Ada>Short form timestamp\n";
    const turns = parseVtt(vtt);
    expect(turns).toEqual([{ speaker: "Ada", startSeconds: 5, text: "Short form timestamp" }]);
  });

  it("falls back to speakerless text when no voice span exists", () => {
    const vtt = "WEBVTT\n\n00:00:01.000 --> 00:00:02.000\nSystem announcement\n";
    const turns = parseVtt(vtt);
    expect(turns).toEqual([{ speaker: "", startSeconds: 1, text: "System announcement" }]);
  });

  it("returns no turns for an empty transcript", () => {
    expect(parseVtt("WEBVTT\n")).toEqual([]);
  });
});

describe("formatTimestamp", () => {
  it("formats minutes and hours", () => {
    expect(formatTimestamp(3)).toBe("0:03");
    expect(formatTimestamp(75)).toBe("1:15");
    expect(formatTimestamp(3723)).toBe("1:02:03");
  });
});

describe("AI insight parsing", () => {
  it("normalizes Graph meetingNotes shapes", () => {
    const notes = parseInsightNotes(
      JSON.stringify([
        { title: "Roadmap", text: "Discussed Q3.", subpoints: [{ title: "Pricing", text: "Needs review" }] },
        "Plain string note",
      ])
    );
    expect(notes).toHaveLength(2);
    expect(notes[0]?.subpoints).toEqual([{ title: "Pricing", text: "Needs review" }]);
    expect(notes[1]?.text).toBe("Plain string note");
  });

  it("normalizes action items and tolerates garbage input", () => {
    expect(parseActionItems(JSON.stringify([{ ownerDisplayName: "Bob", text: "Review pricing" }]))).toEqual([
      { owner: "Bob", text: "Review pricing" },
    ]);
    expect(parseActionItems("not json")).toEqual([]);
    expect(parseInsightNotes("")).toEqual([]);
  });
});

describe("buildPageHtml", () => {
  it("builds a complete page with title, summary on top, transcript below", () => {
    const html = buildPageHtml(
      sampleVtt,
      "Weekly Sync",
      "2026-06-10",
      JSON.stringify([{ title: "Roadmap", text: "Discussed Q3 targets.", subpoints: [] }]),
      JSON.stringify([{ ownerDisplayName: "Bob O'Brien", text: "Review pricing" }])
    );
    expect(html).toContain("<title>2026-06-10 - Weekly Sync</title>");
    expect(html.indexOf("<h1>Summary</h1>")).toBeLessThan(html.indexOf("<h1>Transcript</h1>"));
    expect(html).toContain("<b>Roadmap</b>: Discussed Q3 targets.");
    expect(html).toContain("<b>Bob O&#39;Brien:</b> Review pricing");
    expect(html).toContain("<b>Venkat Yenduri</b> [0:03]");
    // OneNote renders whitespace between tags literally; the page must be compact.
    expect(html).not.toContain("\n");
  });

  it("escapes HTML in meeting subjects and transcript text", () => {
    const vtt = "WEBVTT\n\n00:00:01.000 --> 00:00:02.000\n<v Eve>1 &lt; 2 and <b>bold</b> attempt\n";
    const html = buildPageHtml(vtt, "Q3 <Review> & Planning", "2026-06-10", "[]", "[]");
    expect(html).toContain("<title>2026-06-10 - Q3 &lt;Review&gt; &amp; Planning</title>");
    expect(html).toContain("1 &lt; 2 and bold attempt");
  });

  it("renders a fallback banner when no AI insights exist", () => {
    const html = buildSummaryHtml([], []);
    expect(html).toContain("No AI summary was available");
  });
});
