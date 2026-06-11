// Invoked by .github/workflows/sync.yml on a cron schedule.
// Reads the committed watermark and asks the Power Automate flow to process
// transcripts created inside (since, until]. Fire-and-forget: the standard
// "When a Teams webhook request is received" trigger replies 202 immediately,
// and the flow reports completion back via repository_dispatch.

import { postJson, requireEnv } from "./lib/http.ts";
import { readState } from "./lib/state.ts";
import type { SyncTriggerPayload } from "./lib/types.ts";

// Keeps the Outlook calendar query bounded if the workflow was paused a while.
const MAX_WINDOW_DAYS = 7;

async function main(): Promise<void> {
  const flowUrl = requireEnv("FLOW_URL");
  const secret = requireEnv("FLOW_SECRET");

  const state = readState();
  const until = new Date();
  let since = new Date(state.lastSync);

  const oldest = new Date(until.getTime() - MAX_WINDOW_DAYS * 24 * 60 * 60 * 1000);
  if (since < oldest) {
    console.warn(
      `Watermark ${since.toISOString()} is older than ${MAX_WINDOW_DAYS} days; ` +
        `clamping window start to ${oldest.toISOString()}. Older transcripts will not be synced.`
    );
    since = oldest;
  }
  if (since >= until) {
    console.log(`Watermark ${since.toISOString()} is not in the past; nothing to do.`);
    return;
  }

  const payload: SyncTriggerPayload = {
    since: since.toISOString(),
    until: until.toISOString(),
    secret,
  };

  const status = await postJson(flowUrl, payload);
  console.log(
    `Triggered flow (HTTP ${status}) for window ${payload.since} .. ${payload.until}. ` +
      `The flow will dispatch 'notesync-complete' when done.`
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
