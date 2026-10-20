# Voice Commands for Excel: Project Brief

## Background

This is my fourth personal project. Earlier ones:

1. **AI Investment Research Agent.** A five-agent LangGraph pipeline in Python, with a public web UI and a case study analysing 74 defects from a 95-entry development log. Repo: github.com/patharearya/ai-investment-agent
2. **NC Internship Tracker**
3. **Aqua Static.** A 24-hour hackathon campus app.

I want this one to be different: a new language (TypeScript), a real desktop app instead of a web pipeline, audio input, and a real user to test with.

## The problem

A friend spends a lot of time in desktop Excel on Windows. Selecting many cells, copying columns, formatting, and running calculations takes a lot of clicks, and misclicks keep causing mistakes. He asked whether he could just say what he wants.

Options that exist today are paid (Microsoft 365 Copilot, Dragon) or not built for spreadsheets (Windows Voice Access). This project is a free, spreadsheet-focused voice tool.

## Hard constraint: everything must be free

No paid services, no credit card anywhere. If something needs payment, find a free alternative or drop the feature.

## Habits to carry over from earlier projects

- **Keep a development log (PROJECT_LOG.md) from day one.** One entry per significant defect: what broke, what I first thought had broken, and the measurement that settled it. Read the log first at the start of every session.
- **The LLM interprets, the code acts.** The model only turns speech into a structured command. It never touches the spreadsheet directly. My code validates the command and carries it out.
- **Check claims against the code or the data rather than reasoning about them.** If something can be checked in seconds, check it.
- **A test that cannot fail is not evidence.** Break the thing on purpose and confirm the test goes red.
- **Repeat anything probabilistic before believing the score.** LLM results vary run to run.
- **Measure from the start, not the end.** Build the eval set early.

## What the user sees

A side panel inside desktop Excel, opened from a ribbon button. It stays open while he works.

1. He holds a shortcut key (or clicks the mic) and speaks, e.g. "make column C currency and bold the header row".
2. The panel shows the transcript so mishearings are obvious.
3. The panel shows a preview of exactly what will change ("Format C2:C5 as currency", "Bold A1:D1") and selects those cells in the sheet. Nothing changes yet.
4. He clicks Apply or presses Enter (or says "yes"). Cancel discards it.
5. The last action stays visible at the bottom with an Undo button.

He keeps using mouse and keyboard normally. Voice is for the tedious parts.

## Version 1 scope

**Included**
- Side panel add-in for desktop Excel on Windows
- Push-to-talk (shortcut key and mic button)
- Transcript display
- Preview before apply, Apply/Cancel
- One-step Undo of the last action
- Fixed command menu:
  - **copy**: copy a column or range to another location
  - **format**: bold, italic, number format (currency, percent, number), font color, fill color
  - **compute**: sum, average, min, max, count of a range, written to a target cell
  - **sort**: sort a range by a column, ascending or descending

**Not in version 1** (only if my friend wants more after using it)
- Formulas that refer to columns by header name ("make F price times quantity")
- Trusted mode where simple commands skip the preview
- "Running commands" (undefined, need to ask my friend what he means)
- Multi-step undo, Mac or web support, other languages

## Architecture

```
[Mic in task pane] -> audio blob
      -> [small backend proxy] -> Groq Whisper -> transcript
      -> [small backend proxy] -> Groq LLM (structured output) -> command JSON
      -> [task pane] validate command against schema
      -> preview (select range, list changes)
      -> on Apply: snapshot affected cells, run Office.js, keep snapshot for Undo
```

### Command format

The LLM must return one of a fixed set of command shapes. Validate every response with a schema (for example Zod) before anything runs. Anything that fails validation is shown to the user as "I didn't understand that" with the transcript, and logged. Never guess.

Example shapes:

```json
{ "action": "copy", "from": "C2:C50", "to": "H2" }
{ "action": "format", "range": "A1:D1", "bold": true }
{ "action": "format", "range": "C2:C5", "numberFormat": "currency" }
{ "action": "compute", "op": "sum", "range": "D2:D49", "target": "D50" }
{ "action": "sort", "range": "A1:D50", "by": "C", "order": "desc", "hasHeader": true }
```

A single utterance may produce more than one command (as in the example above), so the response is a list.

Give the LLM the sheet's used range and header row as context, so "the price column" can be resolved to a column letter. Range resolution is the step most likely to go wrong, so the preview exists mainly to catch it.

## Tech stack (all free)

- **Language:** TypeScript
- **Add-in:** Office Add-in with Office.js, scaffolded with Microsoft's Yeoman generator (`yo office`, Excel task pane, TypeScript). `npm start` sideloads it into desktop Excel for development.
- **Speech to text:** Groq Whisper (`whisper-large-v3-turbo`), free tier, no card
- **Command parsing:** a Groq LLM on the free tier, with JSON / structured output
- **Backend proxy:** small Node server (for example Express or Hono). Holds the Groq API key so it is never in the add-in code. Free hosting on Render later; localhost during development.
- **Validation:** Zod
- **Tests:** Vitest for command parsing and validation logic
- **Editor/runtime:** VS Code, Node LTS

## Known technical gotchas

- **Excel's own Ctrl+Z does not undo changes made by an add-in.** Undo must be built: before applying, snapshot the values, formulas, and formats of every affected cell, and restore them on Undo.
- **Microphone access in desktop Excel is the biggest unknown.** The task pane runs in an Edge WebView2 panel. Recording uses `navigator.mediaDevices.getUserMedia` and `MediaRecorder`. Check whether Excel prompts for mic permission and whether the manifest or a device permission API is needed.
- **Do not rely on the browser's built-in Web Speech API.** Embedded browser panels have disabled it before. Use Groq Whisper.
- **Never put the Groq key in the add-in.** Anyone can read add-in code. Calling Groq from the browser may also hit CORS issues. The proxy solves both.
- **Groq free tier limits.** Each transcription request is billed as at least 10 seconds of audio, and there are per-minute request limits. Fine for one user, but add a retry and a clear message when rate limited.
- **Custom keyboard shortcuts** for add-ins on Windows Excel are supported but need the shared runtime setup. Confirm early.
- **Office.js is async and batched.** Work happens inside `Excel.run`, with `context.sync()` to read values. Reading before syncing is a common bug.

## Plan

### Step 0: Spike (one session, before anything else)

Goal: prove the free setup works end to end in desktop Excel. Build the smallest possible add-in with one button that:

1. Records a few seconds of audio from the mic inside the task pane
2. Sends it through the local proxy to Groq Whisper
3. Writes the returned text into cell A1

Answer and log these three questions:

- Can the task pane get microphone access in desktop Excel on Windows?
- Does Groq Whisper on the free tier return a usable transcript?
- Can the add-in write to the sheet?

Also check whether a custom keyboard shortcut can trigger recording. If any answer is no, stop and rethink before building further.

### Step 1: Command parsing, outside Excel

Text in, validated command JSON out. Build and test this as plain TypeScript with no Excel involved, so it can be tested fast. Start the eval set here.

### Step 2: Executing commands in Excel

Implement copy, format, compute, and sort in Office.js, each with a snapshot for Undo.

### Step 3: The full panel

Push-to-talk, transcript, preview, Apply/Cancel, Undo, error states.

### Step 4: Friend testing

Give it to my friend. Watch which commands he actually uses and where it fails. Decide from that whether to stop or add phase-two features.

## Evaluation plan

- Build a set of 50 to 100 spoken commands, ideally recorded by my friend in his own words, each with the expected command JSON.
- Measure separately: transcription accuracy (did Whisper hear it right) and parsing accuracy (did the LLM produce the right command from a correct transcript). Mixing them hides which stage failed.
- Report accuracy by command type (copy, format, compute, sort).
- Keep the recorded audio and transcripts so parsing changes can be re-run against the same inputs without re-recording (frozen-state replay).
- Run each eval more than once and record the variance.

## Open questions

- Do I have desktop Excel on a Windows machine for testing? (Check whether my university provides Microsoft 365.)
- What exactly does my friend mean by "running commands"?
- Folder and repo name for the project.
- Which Groq LLM gives reliable structured output on the free tier? Measure, don't assume.
