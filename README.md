# NoteSync

**After every Teams meeting you organize, a OneNote page appears automatically** — with Copilot's
AI summary and action items at the top, and the full who-said-what transcript below it.

Pages land in a **Meeting Notes** notebook, **Notes** section, one page per meeting, titled
`2026-06-10 - Project Kickoff` style.

## Why it's built this way

It works on a locked-down work account:

- **No IT involvement** — no Entra ID app registration, no client IDs, no admin rights.
- **No premium licenses** — only Standard Power Automate connectors.
- **Your meeting content never leaves Microsoft 365.** GitHub only ever sees timestamps and a
  page count.

## How it works (the short version)

Two halves, talking through a webhook:

1. **GitHub (this repo)** is just the *clock and bookmark*. Every 30 minutes a workflow pings your
   Power Automate flow saying "process meetings from time X to time Y". When the flow finishes, it
   reports back and GitHub saves the new bookmark ([`state/state.json`](state/state.json)) so no
   meeting is ever processed twice.
2. **Power Automate (in your M365 account)** does all the real work: finds your meetings on the
   calendar, fetches each transcript and its Copilot AI summary, turns them into a formatted page
   (using a small TypeScript program that also runs inside M365, as an Office Script), and creates
   the OneNote page.

```
GitHub: "anything new between 9:00 and 9:30?"  ──►  Power Automate: finds meetings,
GitHub: saves bookmark "done through 9:30"     ◄──  builds pages, writes to OneNote
```

## What you need

- An M365 work account **with a Copilot license** (Copilot writes the summaries).
- You must be the **meeting organizer** — meetings you merely attend can't be synced.
- **Transcription turned on** in each meeting (no transcript file means nothing to sync).
- A **private GitHub repo** (your clone of this one).

> **Do I need "AI notes" turned on?** No. Plain transcription is enough — the Copilot summary is
> generated from that same transcript and fetched automatically.

## Setup

One-time, ~45 minutes, no admin help needed:
**[Follow the step-by-step setup guide → docs/SETUP.md](docs/SETUP.md)**

## Good to know

- Pages show up **15–45 minutes after a meeting ends** (transcripts take a few minutes to process,
  and the flow waits 10 extra minutes so Copilot can finish its summary).
- Very short meetings may get a page with no AI summary — Copilot needs enough material to work with.
- If something fails mid-run, the same time window is simply retried on the next tick. In rare
  cases that can create a duplicate page (just delete it).
- The flow's trigger URL acts like a password — if it ever leaks, recreate the trigger and update
  the `FLOW_URL` secret.

## Development

```bash
npm test              # unit tests for the transcript parser / page builder
npm run typecheck     # type-check the whole repo
npm run build:script  # regenerate office-script/transform.ts from src/lib/transform-core.ts
npm run trigger       # what sync.yml runs (needs FLOW_URL / FLOW_SECRET env vars)
npm run complete      # what complete.yml runs (needs GITHUB_EVENT_PATH)
```

`office-script/transform.ts` is **generated** — edit
[`src/lib/transform-core.ts`](src/lib/transform-core.ts) instead, run `npm run build:script`, and
re-paste the output into the Office Scripts editor (Excel → Automate → NoteSync Transform).
