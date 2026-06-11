import { readFileSync, writeFileSync, appendFileSync } from "node:fs";
import { resolve } from "node:path";
import type { SyncState } from "./types.ts";

const STATE_PATH = resolve(process.cwd(), "state", "state.json");
const LOG_PATH = resolve(process.cwd(), "state", "runs.log");

export function readState(): SyncState {
  const raw = readFileSync(STATE_PATH, "utf8");
  const state = JSON.parse(raw) as SyncState;
  if (typeof state.lastSync !== "string" || isNaN(Date.parse(state.lastSync))) {
    throw new Error(`state/state.json has an invalid lastSync: ${raw}`);
  }
  return state;
}

export function writeState(state: SyncState): void {
  writeFileSync(STATE_PATH, JSON.stringify(state, null, 2) + "\n", "utf8");
}

export function appendRunLog(line: string): void {
  appendFileSync(LOG_PATH, line + "\n", "utf8");
}
