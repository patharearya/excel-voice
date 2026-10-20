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
| Task pane gets mic access in desktop Excel? | open | — |
| Groq Whisper free tier returns usable transcript? | open | — |
| Add-in can write to the sheet? | open | — |
| Custom keyboard shortcut triggers recording? | open | — |

Fill the table from the live run, then decide whether to continue.

## Defects
(none yet)
