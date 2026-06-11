// Invoked by .github/workflows/complete.yml when the Power Automate flow
// sends the `notesync-complete` repository_dispatch event. Advances the
// committed watermark so the next cron run starts where this one ended.
// The workflow commits the files this script writes.

import { readFileSync } from "node:fs";
import { appendRunLog, readState, writeState } from "./lib/state.ts";
import { requireEnv } from "./lib/http.ts";
import type { CompletePayload } from "./lib/types.ts";

function readDispatchPayload(): CompletePayload {
  const eventPath = requireEnv("GITHUB_EVENT_PATH");
  const event = JSON.parse(readFileSync(eventPath, "utf8").replace(/^﻿/, "")) as {
    client_payload?: Partial<CompletePayload>;
  };
  const payload = event.client_payload ?? {};
  if (typeof payload.until !== "string" || isNaN(Date.parse(payload.until))) {
    throw new Error(`Dispatch payload has an invalid 'until': ${JSON.stringify(payload)}`);
  }
  return {
    until: payload.until,
    processedCount: typeof payload.processedCount === "number" ? payload.processedCount : -1,
  };
}

function main(): void {
  const payload = readDispatchPayload();
  const state = readState();

  if (Date.parse(payload.until) <= Date.parse(state.lastSync)) {
    console.log(
      `Dispatch watermark ${payload.until} is not newer than ${state.lastSync}; skipping (stale or duplicate event).`
    );
    return;
  }

  writeState({ lastSync: payload.until });
  appendRunLog(
    `${new Date().toISOString()} lastSync=${payload.until} pagesCreated=${payload.processedCount}`
  );
  console.log(
    `Advanced watermark to ${payload.until} (${payload.processedCount} page(s) created).`
  );
}

main();
