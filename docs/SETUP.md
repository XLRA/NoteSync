# Setup guide

Each user runs their own copy of NoteSync: clone the repo (keep it **private**), then work through
the steps below. Everything uses Standard Power Automate connectors — no premium license, no admin
involvement. Budget about 45 minutes.

The flow you'll build does this, in order:

1. Check the request really came from your GitHub repo (shared secret).
2. Ask Outlook for your calendar events in the requested time window.
3. For each Teams meeting you organized: list any new transcripts.
4. For each transcript: fetch its content and the Copilot AI summary.
5. Run an Office Script that turns the transcript into OneNote page HTML.
6. Create the page in OneNote, then tell GitHub the window is done.

## Architecture reference

```
┌─ GitHub (this repo, TypeScript) ─────┐      ┌─ Microsoft 365 (all meeting data stays here) ──────┐
│ sync.yml      every 30 min (cron)    │      │ Power Automate flow "NoteSync" (Standard connectors)│
│   trigger.ts  reads state.json,      │ POST │   1. validate shared secret                         │
│               POSTs {since, until} ──┼─────►│   2. Outlook: calendar events in window             │
│                                      │      │   3. Teams: resolve meeting → list new transcripts  │
│ complete.yml  on repository_dispatch │      │   4. Teams: transcript content + Copilot AI insight │
│   complete.ts commits new watermark ◄┼──────┼─  5. Excel "Run script": office-script/transform.ts │
│               to state/state.json    │ {ts, │      (TypeScript in M365: VTT → OneNote page HTML)  │
│                                      │ count}   6. OneNote: create page in Meeting Notes / Notes  │
└──────────────────────────────────────┘      └─────────────────────────────────────────────────────┘
```

- **Scheduling & state** live in GitHub Actions: the cron workflow tells the flow which time window
  to process; the flow reports back via `repository_dispatch` and the watermark is committed to
  `state/state.json`, so every transcript is processed exactly once.
- **All data processing** happens inside M365. The TypeScript that parses the WEBVTT transcript and
  builds the page HTML runs as an **Office Script** (M365's TypeScript runtime), invoked by the flow
  through the standard Excel Online "Run script" action.
- The flow's entry point is the standard **"When a Teams webhook request is received"** trigger —
  the premium HTTP trigger is not needed.

## Step 1 — One-time M365 prerequisites

Create the places NoteSync writes to and the script it runs, before building the flow.

1. **OneNote**: open OneNote (web is fine) → create a notebook named **Meeting Notes** → inside it
   create a section named **Notes**. Give M365 a few minutes to index it before building the flow
   (the connector dropdowns need to see it).
2. **Excel runner workbook**: in OneDrive, create a folder `NoteSync` and an empty workbook
   `NoteSync/runner.xlsx`. Office Scripts must run "against" a workbook; this one is never written to.
3. **Office Script**: run `npm run build:script` in this repo, then open <https://www.office.com> →
   Excel → open `runner.xlsx` → **Automate** tab → **New Script** → replace the editor contents with
   all of [`../office-script/transform.ts`](../office-script/transform.ts) → rename the script to
   **NoteSync Transform** → **Save script**. To sanity-check it, see
   [Testing the script in isolation](#testing-the-script-in-isolation).
4. **Shared secret**: generate a long random secret, e.g. in PowerShell:
   `-join ((1..48) | ForEach-Object { '{0:x2}' -f (Get-Random -Max 256) })`
   You'll use it in Steps 3 and 8.
5. **Teams**: transcripts only exist for meetings where transcription was on. For meetings you
   organize, open the meeting → **Meeting options** → enable **Record and transcribe automatically**
   (or remember to press **Start transcription**).

## Step 2 — Create the Power Automate flow and trigger

This creates the webhook URL that GitHub will call every 30 minutes.

1. Go to <https://make.powerautomate.com> → **My flows** → **New flow** → **Automated cloud flow**.
2. Name it `NoteSync`, search for trigger **"When a Teams webhook request is received"**
   (Microsoft Teams) → Create.
3. On the trigger card set **Who can trigger the flow?** = **Anyone**. (The URL contains an
   unguessable signature, and Step 3 adds a shared-secret check on top.)
4. **Save** the flow once — the trigger card then shows the **HTTP POST URL**. Copy it for Step 8.

The GitHub workflow POSTs this JSON body, available in the flow as `triggerBody()`:

```json
{ "since": "2026-06-10T00:00:00.000Z", "until": "2026-06-10T18:00:00.000Z", "secret": "..." }
```

## Step 3 — Validate the secret

Anyone who somehow got the trigger URL still can't run the flow without the secret.

1. Add a **Condition** action named `Check secret`:
   - left: expression `triggerBody()?['secret']`
   - **is equal to**
   - right: paste the secret from Step 1.4 (it lives inside the flow definition; only flow
     co-owners can see it)
2. In the **If no** branch, add **Terminate** → Status: **Cancelled**.
3. All remaining steps go in the **If yes** branch.

## Step 4 — Find your Teams meetings in the window

Pull the calendar for the requested window, keep only Teams meetings you organized, and reduce each
one to its join URL (which is how the Teams connector looks meetings up).

1. **Initialize variable** → name `processedCount`, type **Integer**, value `0`.
   Add two more **string** variables here: `notesJson` = `[]` and `actionsJson` = `[]` (used in Step 6).
2. **Office 365 Outlook → Get calendar view of events (V3)**:
   - Calendar id: `Calendar`
   - Start time: expression `addHours(triggerBody()?['since'], -24)`
     *(24h of slack so long meetings whose transcript appears late are still found)*
   - End time: expression `triggerBody()?['until']`
3. **Filter array** named `Organized Teams meetings`:
   - From: `outputs('Get_calendar_view_of_events_(V3)')?['body/value']`
   - Switch to **Edit in advanced mode** and paste:
     ```
     @and(equals(item()?['isOrganizer'], true), contains(coalesce(item()?['body'], ''), 'https://teams.microsoft.com/l/meetup-join/'))
     ```
4. **Select** named `Join URLs` (maps each event to just its Teams join link):
   - From: the `Filter array` output (`body('Organized_Teams_meetings')`)
   - Toggle the Map field to **text mode** (T icon) and paste this expression, which extracts the
     join link from the invite body:
     ```
     concat('https://teams.microsoft.com/l/meetup-join/', first(split(first(split(last(split(item()?['body'], 'https://teams.microsoft.com/l/meetup-join/')), '"')), '''')))
     ```
5. **Compose** named `Unique join URLs` — dedupes recurring-meeting instances (every occurrence
   shares one join URL):
   ```
   union(body('Join_URLs'), body('Join_URLs'))
   ```

> Note: if your events expose **Online meeting URL** directly (some tenants do), you can use that
> field instead of the body-splitting expression in step 4 — the result must be the full
> `https://teams.microsoft.com/l/meetup-join/...` link.

## Step 5 — Per meeting: resolve ID and list new transcripts

Loop over the meetings, turn each join URL into a meeting ID, and find transcripts that haven't
been processed yet.

1. **Apply to each** over `outputs('Unique_join_URLs')`. Open the loop's **Settings** and set
   **Concurrency control → On, Degree of parallelism = 1** (keeps `processedCount` accurate).
2. Inside the loop, add a **Scope** named `Try meeting`, and put steps 3–4 plus all of Step 6 inside
   it. After the scope, add a **Compose** named `Meeting skipped` (any placeholder text) and set its
   **Configure run after** to run only when `Try meeting` **has failed** or **is skipped** — this
   stops one odd meeting (cancelled, lookup failure) from blocking every later run. Then add a second
   **Compose** named `Continue` configured to run after `Meeting skipped` **is successful** or
   **is skipped**, so the flow always proceeds.
3. **Microsoft Teams → Get an online meeting**:
   - Lookup: **Join web URL** = `item()` (the current join URL)
4. **Microsoft Teams → List meeting transcripts**:
   - Meeting ID: `outputs('Get_an_online_meeting')?['body/id']`
5. **Filter array** named `New transcripts`:
   - From: the transcripts list output
   - Advanced mode:
     ```
     @and(greater(item()?['createdDateTime'], triggerBody()?['since']), lessOrEquals(item()?['createdDateTime'], triggerBody()?['until']), less(item()?['createdDateTime'], addMinutes(triggerBody()?['until'], -10)))
     ```
   The first two clauses are the dedup watermark (each transcript is processed exactly once across
   runs). The third skips transcripts younger than 10 minutes so Copilot has time to generate AI
   insights — they're picked up by the next half-hourly run instead.

## Step 6 — Per transcript: fetch, transform, write to OneNote

This is the payoff: grab the transcript and the Copilot summary, run the Office Script to build the
page HTML, and create the OneNote page.

Inside a nested **Apply to each** over `body('New_transcripts')` (also set parallelism = 1):

1. **Microsoft Teams → Get meeting transcript content**:
   - Meeting ID: `outputs('Get_an_online_meeting')?['body/id']`
   - Transcript ID: `item()?['id']`
2. **Microsoft Teams → List AI insights** — Meeting ID as above.
3. **Condition** `Has AI insights`: expression `empty(outputs('List_AI_insights')?['body/value'])`
   **is equal to** `false`.
   - **If yes**: **Microsoft Teams → Get AI insight** (Meeting ID as above, AI Insight ID:
     `last(outputs('List_AI_insights')?['body/value'])?['id']`), then **Set variable** `notesJson` =
     `string(coalesce(outputs('Get_AI_insight')?['body/meetingNotes'], json('[]')))` and
     **Set variable** `actionsJson` = `string(coalesce(outputs('Get_AI_insight')?['body/actionItems'], json('[]')))`.
   - **If no**: **Set variable** `notesJson` = `[]` and **Set variable** `actionsJson` = `[]`.
     (Happens when a meeting was too short to produce insights — the page gets a "no summary" note.)
4. **Excel Online (Business) → Run script**:
   - Location: **OneDrive for Business**, Document Library: **OneDrive**, File: `/NoteSync/runner.xlsx`
   - Script: **NoteSync Transform**
   - Parameters (these appear once the script is selected — order/names must match):
     - `vtt`: the transcript content from step 1 (dynamic content; if it offers a file/binary token
       use expression `string(body('Get_meeting_transcript_content'))`)
     - `subject`: `coalesce(outputs('Get_an_online_meeting')?['body/subject'], 'Meeting')`
     - `dateStr`: `formatDateTime(convertTimeZone(item()?['createdDateTime'], 'UTC', 'Eastern Standard Time'), 'yyyy-MM-dd')`
       *(swap in your Windows time-zone name)*
     - `notesJson`: `variables('notesJson')`
     - `actionItemsJson`: `variables('actionsJson')`
5. **OneNote (Business) → Create page in a section**:
   - Notebook Key: **Meeting Notes** (pick from dropdown)
   - Notebook section: **Notes** (pick from dropdown)
   - Page Content: the Run script **result** (dynamic content / `outputs('Run_script')?['body/result']`)
   The page title comes from the `<title>` tag the script generates: `yyyy-MM-dd - Meeting Name`.
6. **Increment variable** `processedCount` by 1.

## Step 7 — Report completion back to GitHub

Tell GitHub the window succeeded so it advances the bookmark and never re-sends this window.

After (outside) the outer Apply to each, still in the **If yes** branch:

1. **GitHub → Create a repository dispatch event (Preview)**. Sign in with your GitHub account when
   prompted to create the connection.
   - Repository owner: your GitHub username
   - Repository name: `NoteSync`
   - Event type: `notesync-complete`
   - Client payload (switch to expression/JSON input):
     ```json
     { "until": "@{triggerBody()?['until']}", "processedCount": "@{variables('processedCount')}" }
     ```

If anything upstream fails, the flow run fails and this dispatch never fires — the GitHub watermark
stays put and the same window is retried on the next cron tick. That makes the pipeline
at-least-once; the per-transcript window filter in Step 5 keeps retries from double-posting pages
that already succeeded in a *previous completed* run. (A duplicate page is possible only if a run
dies *between* creating a page and dispatching — rare; delete the duplicate by hand or add a
"Get pages for a specific section" title check before Step 6.5 if it ever bothers you.)

## Step 8 — GitHub repo configuration

Give the GitHub workflow the flow's address and the shared secret.

In GitHub → your repo → Settings → Secrets and variables → Actions, add:

- `FLOW_URL` = the HTTP POST URL from Step 2.4 (treat it like a password)
- `FLOW_SECRET` = the secret from Step 1.4

Then set [`../state/state.json`](../state/state.json) `lastSync` to the current time and adjust the
cron hours in [`../.github/workflows/sync.yml`](../.github/workflows/sync.yml) to your working
hours (UTC).

## Step 9 — End-to-end test

1. Create a Teams meeting with yourself (+ one colleague or a second account), start it,
   **Start transcription**, talk for a minute, end it.
2. Wait ~15 minutes (transcript processing + the flow's 10-minute insight buffer).
3. GitHub → Actions → **Sync transcripts** → **Run workflow**.
4. Watch: the flow run in <https://make.powerautomate.com> → My flows → NoteSync → 28-day run
   history; then the OneNote page in **Meeting Notes → Notes**; then the **Record sync completion**
   workflow run and the `state/state.json` commit it pushes.

## Testing the script in isolation

In the Office Scripts editor you can't pass parameters from the Run button, so add this temporarily
at the bottom, run it, check the console output, then remove it:

```ts
function test(workbook: ExcelScript.Workbook) {
  const vtt = 'WEBVTT\n\n00:00:01.000 --> 00:00:02.000\n<v Test User>Hello world</v>\n';
  console.log(main(workbook, vtt, "Test Meeting", "2026-06-10", "[]", "[]"));
}
```

## Troubleshooting

| Symptom | Likely cause |
| --- | --- |
| Flow runs but terminates immediately | Secret mismatch between `FLOW_SECRET` and the Condition value |
| `Get an online meeting` 404s | You weren't the organizer, or the join URL extraction grabbed a truncated link — inspect the `Join URLs` output |
| `List meeting transcripts` is always empty | Transcription wasn't turned on in the meeting, or admin policy blocks transcription |
| AI insight branch always empty | Copilot disabled for the meeting, or the meeting was too short for insights (they can take up to 4 hours for long meetings) |
| OneNote action can't see the notebook | Notebook not in the default OneDrive Notebooks location, or not indexed yet — wait and re-add the action |
| Page renders with stray blank lines | Page content was edited to have whitespace between HTML tags — pass the script result through untouched |
| `complete.yml` never runs | GitHub connection in the flow lost authorization, or dispatch event type isn't exactly `notesync-complete` |
