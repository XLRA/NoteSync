/** Body POSTed to the flow's "When a Teams webhook request is received" trigger. */
export interface SyncTriggerPayload {
  /** Process transcripts created strictly after this ISO timestamp. */
  since: string;
  /** ...and up to (inclusive) this ISO timestamp. */
  until: string;
  /** Shared secret validated by the flow before doing any work. */
  secret: string;
}

/** client_payload of the `notesync-complete` repository_dispatch sent by the flow. */
export interface CompletePayload {
  /** The window end the flow finished processing; becomes the new watermark. */
  until: string;
  /** Number of OneNote pages created in this run. */
  processedCount: number;
}

export interface SyncState {
  /** Watermark: transcripts created at or before this instant are done. */
  lastSync: string;
}
