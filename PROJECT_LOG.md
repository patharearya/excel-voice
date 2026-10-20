# Project log

Read this first at the start of every session. One entry per significant
defect: what broke, what I first thought had broken, the measurement that
settled it.

## 2026-10-19 — Day 0

### Environment facts (checked, not assumed)
- Desktop Excel is installed: `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`. Answers open question 1 (yes, Windows desktop Excel is available).
- Node v24.15.0, npm 11.12.1, git 2.54.
- No `GROQ_API_KEY` in the environment yet. Needs a free Groq account (console.groq.com, no card).

### Step 0 spike status
Built: add-in with one Record button + localhost proxy to Groq Whisper.
Not yet run live. The three spike questions are still OPEN until someone
clicks the button inside Excel:

| Question | Answer | Evidence |
|---|---|---|
| Task pane gets mic access in desktop Excel? | yes | run 4: 94 KB, peak 100%, real transcript (Defect 1 was an idle Bluetooth mic) |
| Groq Whisper free tier returns usable transcript? | yes (TTS and real mic) | proxy log 2026-10-20T03:00:06Z, 200 in 720 ms, word-exact |
| Add-in can write to the sheet? | yes | "Thank you." written to A1 |
| Custom keyboard shortcut triggers recording? | no (cause unconfirmed) | getShortcuts() returns {} across 4 manifest variants, see Defect 2 |

Fill the table from the live run, then decide whether to continue.

## Defects
(none yet)

### Doc findings for the spike (2026-10-19)
- **Mic permission.** Microsoft's message-centre notice MC792499 (May 2024) says
  "Add-ins that run on Office on Windows, Mac, or Safari already have prompts in
  place"; the `Office.devicePermission.requestPermissions` API is only for the
  web / Chromium hosts. So on desktop Windows the expectation is: call
  `getUserMedia`, Excel/WebView2 shows its own prompt. No manifest change
  needed. Still unverified until the live run.
- **Keyboard shortcut.** Needs the shared runtime (`SharedRuntime 1.1` in
  `<Requirements>`, `<Runtimes><Runtime lifetime="long">`, FunctionFile =
  Taskpane.Url) plus `<ExtendedOverrides Url=".../shortcuts.json">` after
  `</VersionOverrides>` and `Office.actions.associate(id, fn)`. Supported on
  Excel for Windows from Version 2102. Recommended key pattern
  Ctrl+Shift+Alt+x; chose Ctrl+Shift+Alt+V. Done in `addin/manifest.xml` and
  `addin/assets/shortcuts.json`.
  Sources: learn.microsoft.com/office/dev/add-ins/design/keyboard-shortcuts,
  .../develop/configure-your-add-in-to-use-a-shared-runtime
- **Proxy self-check** (no Groq key available): missing key -> exit 1; OPTIONS
  -> 204; GET / -> 404; POST with a bad key -> Groq's own 401 JSON passed
  through in 240 ms. The proxy reaches Groq; only the key is missing.
- **Risk to watch:** task pane is https://localhost:3000, proxy is
  http://localhost:8787. Chromium treats localhost as secure so mixed-content
  blocking should not apply, but WebView2 is the thing under test.

### 2026-10-19 live run, part 1 (proxy only, no Excel yet)
- Key valid: GET /openai/v1/models -> 200, lists whisper-large-v3-turbo,
  openai/gpt-oss-120b, openai/gpt-oss-20b, qwen/qwen3.8-27b.
- Synthetic test: Windows TTS WAV of "Make column C currency and bold the
  header row" (149 KB, 16-bit PCM) posted to the proxy.
  Groq 200 in 720 ms, text: " Make column C currency and bold the header row."
  Word-exact. **Question 2 (usable transcript on free tier): YES.**
  Caveat: TTS audio is cleaner than a real mic; the mic run still matters.
- Defect: the key was first pasted into `proxy/.env.example`, which is
  tracked. Moved to `proxy/.env` (ignored), placeholder restored. If that
  key is ever committed anywhere, rotate it at console.groq.com.
- `npm start` in addin/ printed "Unable to start the dev server ... not running
  on port 3000" then launched Excel anyway. First thought: cert not installed.
  Measured: cert already trusted, and 30 s later port 3000 answered 200 for
  taskpane.html and assets/shortcuts.json. The dev server just came up slower
  than the launcher's wait. Harmless; if the pane is blank, close and reopen it.

### 2026-10-19 live run, part 2 (in Excel) — DEFECT 1: "Thank you."
**What broke.** Said "Copy cell A2 into cell B3"; pane showed "Thank you."
Two attempts, both "Thank you.". Shortcut Ctrl+Shift+Alt+V: user reports
nothing happened.
**First thought.** Whisper misheard.
**Measurement.** Proxy log: uploads were 1696 and 1892 bytes of
audio/webm;codecs=opus. A few seconds of speech in Opus is 10-30 KB; under
2 KB is a header plus silence. "Thank you." is Whisper's known output for
silent audio. So the mic path ran end to end (getUserMedia resolved,
MediaRecorder produced a blob, proxy -> Groq 200, A1 was written) but the
audio contained no speech. Not a Whisper problem.
**Ruled out.** Windows privacy: microphone global=Allow, desktop apps
(NonPackaged)=Allow. Excel 16.0.20430, above the 2102 shortcut minimum.
**Open.** Why silent. Candidates: wrong default input device (three
present: Intel SST Bluetooth, Intel SST Digital Microphones, Realtek), track
muted, or WebView2 delivering a silent stream. Note the ConsentStore lists
msedge, WINWORD, POWERPNT as having requested the mic, but not EXCEL.
**Instrumented for the next run.** Pane now shows the live input level, the
track label/muted flag/settings, duration and peak level. Proxy saves every
upload to proxy/recordings/ for replay. Shortcut handler logs "shortcut fired".

Spike table so far: mic access = yes (no error, a stream was returned);
transcript = yes but audio silent in Excel; sheet write = yes ("Thank you."
landed in A1); shortcut = open.
- Run 3 with level meter: 1766 bytes, 6.4 s, **peak level 0%**, no bar
  movement. The stream is silent, not quiet. Confirms the diagnosis above.
- Audio endpoints (Get-PnpDevice AudioEndpoint): "Arya's AirPods #2
  Hands-Free" headset is connected (status OK) alongside the Intel Microphone
  Array. Leading hypothesis: the Bluetooth hands-free capture endpoint is the
  default mic and delivers silence while the AirPods are idle/in the case.
  Next: pane now lists every audio input and lets you pick one; the mic line
  (device label, muted flag) stays on screen after Stop.
- Excel status bar shows "Error installing add-ins". OfficeAddins.log.txt has
  nothing beyond "manifest parsing starting" (5 lines). Possibly the
  ExtendedOverrides (shortcuts.json) fetch; unconfirmed. Shortcut still open.
- Tooling defect: a Node patch script written through the Bash tool had its
  `\n` collapsed to `\n` (see CLAUDE.md), so the anchor did not match. Wrote
  the file with the Write tool instead.
- Run 4: with the device list visible, the default input was
  "Default - Headset (AirPods) (Bluetooth)" (5 inputs listed). This time
  93998 bytes, 6.6 s, peak 100%, transcript "Copy cell A2 and to cell C3."
  Same default device as the silent runs, so the AirPods were simply idle
  earlier. **Defect 1 resolved: silent Bluetooth hands-free endpoint, not the
  add-in.** For the friend's machine: show the device picker, and warn when
  peak level stays at 0%.
- Spike verdict: mic access YES, Whisper YES, sheet write YES. Shortcut still
  open; "Error installing add-ins" still unexplained.
- Run 5: "Copy cell B3 into cell D5." word-exact (116 KB, 8.3 s). Run 4 was
  also word-exact per the user. Eval set entries 1 and 2.

### DEFECT 2: shortcut never fires
**What broke.** Ctrl+Shift+Alt+V does nothing. Pane diag line:
`SharedRuntime 1.1: true | KeyboardShortcuts 1.1: true | Office.addin: object | registered shortcuts: {}`.
Excel status bar: "Error installing add-ins".
**First thought.** ExtendedOverrides URL not fetched (localhost cert).
**Measurement.** Excel's cache has the file:
`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{...}\ExtendedManifest\<id>_1.0.0.0_en-US`,
byte-identical to ours including the `$schema` key. So fetch and cert are
fine; Excel parsed it and registered nothing. Microsoft's sample
shortcuts.json has no `$schema` key and its manifest has no `<Requirements>`
block. Change one thing: removed `$schema`, bumped Version to 1.0.0.1 so
the cache key changes, re-sideloaded. Pending result.
- Removing `$schema` + version bump: Excel cached the new file
  (`..._1.0.0.1_en-US`, 217 bytes) and still `registered shortcuts: {}`.
  Not the cause.
- Control: sideloaded Microsoft's office-keyboard-shortcuts sample (hosted on
  GitHub Pages, id 05c2e1c9-...) via the WEF\Developer registry key beside
  ours. Its Ctrl+Alt+Up should show its task pane. If the sample's shortcut
  also fails, the defect is this Excel/account, not our manifest. Only
  remaining manifest difference: our `<Requirements>` block (sample has none).
- Control inconclusive: Excel cached the sample's manifest (Wef\Manifests)
  but never fetched its shortcuts.json and never showed its ribbon button.
  Runtime log (7 lines) records no error. Unknown why; registry sideload of
  a second dev add-in may need the Excel "Upload My Add-in" route instead.
- Run 6 (mic, device "Communications - Headset (AirPods)"): "Copy cell B2
  into cell D3." word-exact, 97 KB, 6.7 s. Eval entry 3. Mic path is solid.
- Last manifest difference removed: `<Requirements>` block deleted, Version
  1.0.0.2, re-sideloaded. Manifest still validates. Pending result.
- Decision rule: if `registered shortcuts` is still `{}` after this, mark the
  shortcut question "no (cause unconfirmed)", ship the mic button as the v1
  trigger, and revisit on the friend's machine. Three cheap leads left:
  Excel > Add-ins > More Add-ins > Developer Add-ins (does the sample list?),
  upload the sample via "Upload My Add-in", and check KeyboardShortcuts on
  Excel on the web with the same manifest.
- Without `<Requirements>` (1.0.0.2): still `registered shortcuts: {}`.
  Four variants tried (with/without `$schema`, with/without `<Requirements>`),
  each confirmed fetched and cached by Excel. **Defect 2 parked, cause
  unconfirmed.** Sample add-in removed from the registry.

### Spike verdict (2026-10-19, end of session)
| Question | Answer |
|---|---|
| Mic in task pane | Yes. Pick the device; an idle Bluetooth headset records silence. |
| Groq Whisper free tier | Yes. 3/3 real recordings word-exact, under 1 s each. |
| Write to sheet | Yes. |
| Keyboard shortcut | No, cause unconfirmed. Mic button is the v1 trigger. |

Proceed to Step 1 (command parsing outside Excel). Eval set seeds, all from
the AirPods mic in Excel:
1. "Copy cell A2 and to cell C3." (user confirmed verbatim)
2. "Copy cell B3 into cell D5."
3. "Copy cell B2 into cell D3."
Recordings for 2 and 3 are in proxy/recordings/.

Running processes to stop when done: `npm run stop` in addin/ (dev server +
sideload), and the proxy (Ctrl+C or kill the node on port 8787).
- Final shortcut attempt: Excel closed, whole Wef cache (521 MB) deleted,
  fresh sideload of 1.0.0.2, shortcuts.json refetched. Still
  `registered shortcuts: {}`. Closed as "no, cause unconfirmed" for this
  machine. Run 7: "Copy cell B2 into cell C4." word-exact, 174 KB, 11.8 s.
  Eval entry 4.

## Step 1: command parsing outside Excel (started 2026-10-19 late)
Lives in `proxy/` so the Groq key stays server-side: Zod schema is the single
source of truth (also emitted as the JSON schema sent to Groq), `interpret`
calls the LLM, an eval script replays a fixed case list N times per model.
- Files: proxy/commands.ts (Zod, nullable-not-optional so Groq strict mode
  accepts the same schema), interpret.ts (prompt states the range-resolution
  mechanism; temperature 0; strict json_schema; 429 retry), eval.ts,
  eval/cases.json (30 cases: 4 mic transcripts from tonight + 26 synthetic,
  4 of them "none" cases including "Thank you." from silent audio),
  commands.test.ts (13 Vitest checks).
- Test-can-fail check: replaced the range regex with /^.*$/ -> 2 of 13 tests
  red (lowercase range, sheet-qualified range). Restored -> 13 green.
- First live call, gpt-oss-120b, "make column C currency and bold the header
  row" -> exactly the two commands from the brief, 1053 ms, 277 reasoning
  tokens. Strict mode accepted Zod 4's toJSONSchema output unchanged.

### DEFECT 3: eval runner crashed; first eval results
**What broke.** gpt-oss-120b run died after ~27 cases: uncaught
`ConnectTimeoutError` from fetch. No 120b result saved.
**First thought.** Groq outage.
**Measurement.** The 20b run right after completed fine, so it was a
transient connect timeout. The instrument was wrong: eval.ts and interpret.ts
had no handling for a thrown fetch. Fixed: network errors retry once and
otherwise come back as a rejection with reason `network: <code>`, so a row is
recorded instead of the run dying.
**Also found.** Nearly every call hit 429 (free tier ~30 req/min). Retries
worked but inflated latency (median 8.8 s). Now: 2 s gap between eval calls,
and `ms` measures only the final attempt.

**gpt-oss-20b, 30 cases x 3 runs (before the fixes; scores unaffected):**
| type | run1 | run2 | run3 | all |
|---|---|---|---|---|
| copy | 63% | 75% | 63% | 67% |
| format | 100% | 100% | 100% | 100% |
| compute | 100% | 100% | 100% | 100% |
| sort | 100% | 100% | 100% | 100% |
| none | 100% | 100% | 100% | 100% |
| total | 90% | 93% | 90% | 91% |
Every failure is one shape: whole-column copy ("copy column C to column H")
returned C2:C50 -> H2, expected C1:C50 -> H1. The prompt states the
header-inclusion rule for copy and 20b ignores it consistently (copy-08 flaky).
All 4 real mic transcripts passed 3/3. All 4 "none" cases returned [] 3/3.
