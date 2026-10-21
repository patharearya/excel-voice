# Excel Voice

Talk to Excel. Say *"make column C currency and bold the header row"*, check the
preview, press Enter, and it's done.

Excel Voice is a free side panel for desktop Excel on Windows. It is for anyone
who spends a long time selecting cells, formatting and copying columns, and
would rather say what they want than click through it.

- **Free.** It uses Groq's free tier for speech recognition and understanding.
  No credit card anywhere.
- **Nothing changes until you approve it.** Every command is shown as a
  preview with the affected cells selected. You press Apply or Cancel.
- **One-step Undo.** Excel's own Ctrl+Z can't undo changes made by an add-in,
  so Excel Voice keeps its own copy of the cells it changes.

> **Status:** early, working prototype (personal project). It runs on your own
> computer: you start it from two Command Prompt windows, then use the panel
> inside Excel. There is no one-click installer yet, but the steps below
> assume no coding experience.

---

## What you can say

| Kind | Examples |
|---|---|
| **Format** | "Bold the header row", "Make the price column currency", "Highlight B2 to B10 yellow", "Make the total column red", "Format D2 to D50 as a percent" |
| **Calculate** | "Total column D at the bottom", "Average the price column into C51", "Put the max of column B in F1" (writes a live formula: SUM, AVERAGE, MIN, MAX or COUNT) |
| **Change numbers** | "Double column B", "Add 10 to C2 through C20", "Increase the prices by 10 percent", "Halve D10" |
| **Sort** | "Sort by price, high to low", "Sort the table by quantity ascending" |
| **Copy** | "Copy column C to column H", "Copy A2 through A10 to B2" |

You can combine commands in one sentence ("make column C currency and bold the
header row"). Columns can be named by letter or by their header ("the price
column"), as long as the headers are in row 1.

Anything outside this list ("delete column B", "make F price times quantity")
gets *"I didn't catch a command in that"*. Excel Voice never guesses.

---

## What you need

- **A Windows computer with desktop Excel** (Microsoft 365, or Office 2021 or
  later). Excel in a web browser and Excel on Mac won't work.
- **A microphone.** Your laptop's built-in one is fine. You can also type
  commands instead of speaking.
- **About 20 minutes** for the one-time setup below.

You don't need to know how to code. You will copy and paste a few lines into a
window called **Command Prompt**; every step below says exactly what to paste
and what you should see.

---

## Part 1: One-time setup

Do Part 1 once. After that, you only need Part 2 each time you use it.

### Step 1. Install Node.js

Node.js is a free program that runs Excel Voice's two small helper programs.

1. Go to <https://nodejs.org>.
2. Click the big button that says **LTS** (it downloads a file ending in
   `.msi`).
3. Open the downloaded file and click **Next** on every screen, then
   **Install**, then **Finish**. The default options are all fine.

### Step 2. Install Git

Git is a free program that downloads this project for you.

1. Go to <https://git-scm.com/download/win>.
2. Click **Click here to download** (the 64-bit version).
3. Open the downloaded file and click **Next** on every screen (there are a
   lot; the defaults are fine), then **Install**, then **Finish**.

### Step 3. Open Command Prompt

Command Prompt is a window where you type (or paste) instructions for your
computer, one line at a time.

1. Press the **Windows key** on your keyboard.
2. Type **cmd**.
3. Click **Command Prompt** in the results.

A black window opens with a line like:

```
C:\Users\YourName>
```

This is where you'll paste every command in Part 1. **To run a command:**
copy it from this page, click inside the black window, press **Ctrl+V** to
paste it (or right-click), then press **Enter**. Wait until the
`C:\Users\YourName>` line appears again before pasting the next one.

**Check that steps 1 and 2 worked.** Paste this and press Enter:

```
node --version
```

You should see a version number such as `v24.11.0`. Then paste this and press
Enter:

```
git --version
```

You should see something like `git version 2.51.0`.

> If you see **"is not recognized as an internal or external command"**,
> close the Command Prompt window, open a new one (Step 3 again), and retry.
> Programs installed while Command Prompt was open aren't visible to it until
> you reopen it. If it still fails, repeat Step 1 or Step 2.

### Step 4. Download Excel Voice

In the same Command Prompt window, paste this and press Enter:

```
git clone https://github.com/patharearya/excel-voice.git
```

You'll see a few lines ending in `done.` This created a folder called
**excel-voice** inside your user folder (`C:\Users\YourName\excel-voice`).

### Step 5. Get your free Groq key

Groq is the free online service that turns your speech into text and works
out what you meant. It needs a "key", a long password that identifies you.

1. Go to <https://console.groq.com> and sign up (Google account or email).
   It doesn't ask for a credit card.
2. Go to <https://console.groq.com/keys> and click **Create API Key**.
3. Type any name (for example `excel`) and click **Submit**.
4. Click the copy button next to the key. It starts with `gsk_`.
   **Keep this page open**: Groq only shows the key once. If you lose it,
   just create another.

### Step 6. Set up the voice service

Back in Command Prompt, paste each line below **one at a time**, pressing
Enter after each and waiting for it to finish.

Go into the project's `proxy` folder:

```
cd excel-voice\proxy
```

The line now ends in `\excel-voice\proxy>`. Download what it needs (takes
under a minute):

```
npm install
```

It finishes with a line like `added 36 packages` (a note about packages
"looking for funding" is normal too). Lines starting with `npm warn` are fine;
only lines starting with `npm error` mean something went wrong.

Make your settings file:

```
copy .env.example .env
```

You should see `1 file(s) copied.` Now open it:

```
notepad .env
```

Notepad opens with one line:

```
GROQ_API_KEY=gsk_...
```

Delete `gsk_...` and paste your key in its place, so the line reads
`GROQ_API_KEY=gsk_` followed by the rest of your key, with no spaces. Click
**File → Save**, then close Notepad.

> Your key stays in this file on your computer. Don't share it or post it
> anywhere. If it ever leaks, delete it on the Groq keys page and make a new
> one.

### Step 7. Set up the Excel add-in

Still in Command Prompt, paste these one at a time:

```
cd ..\addin
```

The line now ends in `\excel-voice\addin>`.

```
npm install
```

This one takes a few minutes and prints a lot. Wait for the
`C:\Users\YourName\excel-voice\addin>` line to come back.

**Part 1 is done.** You can close Command Prompt.

---

## Part 2: Start Excel Voice (every time)

Excel Voice needs **two Command Prompt windows** open while you use it: one
runs the voice service, the other loads the panel into Excel.

**First, close Excel** if it's open. Excel Voice opens it for you.

### Window 1: start the voice service

1. Open Command Prompt (Windows key, type **cmd**, click **Command Prompt**).
2. Paste these one at a time, pressing Enter after each:

   ```
   cd excel-voice\proxy
   ```

   ```
   npm start
   ```

3. You should see:

   ```
   proxy listening on http://localhost:8787 ...
   ```

**Leave this window open**, and don't type anything else in it. You can
minimise it. If you close it, the panel can't hear you.

### Window 2: open Excel with the panel

1. Open a **second** Command Prompt window the same way.
2. Paste these one at a time:

   ```
   cd excel-voice\addin
   ```

   ```
   npm start
   ```

3. **The first time only**, Windows shows pop-ups about installing a
   **certificate** ("Developer CA for Microsoft Office Add-ins"), and possibly
   "Do you want to allow this app to make changes to your device?". Click
   **Yes** on each. This lets the panel run securely on your own computer.
4. **A third black window opens by itself** and fills with text. That's the
   panel's helper. Leave it open too (minimising is fine).
5. Excel opens after 10–30 seconds. Window 2 might say "Unable to start the
   dev server" first; that only means the helper was slow to start, and Excel
   opens anyway.
6. In Excel, open a workbook (a blank one is fine). On the **Home** tab,
   look at the far right end of the ribbon for **Excel Voice** (in a group
   called **Voice**) and click it. The panel opens on the right side.

> **Blank panel?** It opened before everything was ready. Close the panel
> with the **X** in its top corner and click **Excel Voice** again.

---

## Use it

1. Click **Speak**, say your command, click **Stop**. (Or type the command
   into the box under the button and press **Enter**.)
2. The panel shows what it heard, a list under **Will change**, and selects
   those cells in your sheet. Nothing has changed yet.
3. Check the selection. Press **Enter** (or **Apply**) to do it, **Esc** (or
   **Cancel**) to drop it.
4. The change appears under **Last change** with an **Undo** button. Undo
   puts the cells back exactly as they were: values, formulas and formatting.

Not sure what to say? With nothing pending, the panel shows example commands.
Click one to try it.

**First recording:** Windows or Excel may ask for permission to use the
microphone. Click **Allow**.

---

## Stop it

1. Click into **Window 2** (the one you used to open Excel) and paste:

   ```
   npm run stop
   ```

   This removes the panel from Excel and closes the helper window that opened
   by itself.
2. Click into **Window 1** (the voice service) and press **Ctrl+C**. If it
   asks `Terminate batch job (Y/N)?`, type `Y` and press Enter.
3. Close both Command Prompt windows, and close Excel as normal.

Next time, start again from **Part 2**.

---

## Troubleshooting

| What you see | What to do |
|---|---|
| **"… is not recognized as an internal or external command"** (for `node`, `npm` or `git`) | Close Command Prompt and open a new one. If it still happens, install Node.js or Git again (Part 1, Steps 1–2). |
| **"The system cannot find the path specified"** after `cd excel-voice\...` | Command Prompt isn't in your user folder. Close it, open a new one (it starts in the right place), and paste the line again. If it still fails, Step 4 (download) didn't finish; run it again. |
| **"running scripts is disabled on this system"** | You're in PowerShell, not Command Prompt. Open **Command Prompt** instead (Windows key, type **cmd**). |
| **Window 1 shows "GROQ_API_KEY is not set"** | The key wasn't saved. Repeat the `notepad .env` part of Step 6, make sure you clicked **File → Save**. |
| **No Excel Voice button on the Home tab** | Check Window 2 for an error. Run `npm run stop` there, then `npm start` again. |
| **"No sound was picked up"** | The mic recorded silence. Open **Microphone** at the bottom of the panel and pick another device. Bluetooth headsets (AirPods, etc.) often record silence until they're actively in use. |
| **"Microphone access is blocked"** | Windows Settings → Privacy & security → Microphone: turn on microphone access, and access for desktop apps. |
| **"Can't reach the voice service"** | The voice service isn't running. Start it (Part 2, Window 1) and keep that window open. |
| **"Too many requests right now"** | Groq's free tier allows about 30 requests a minute. Wait a few seconds. |
| **"I didn't catch a command in that"** | The sentence wasn't one of the supported commands, or it was misheard. Check the quoted transcript in the panel, then rephrase or type it. |
| **"Can't do this safely: …"** | The command was understood but would cause a problem, such as putting a total inside the range it adds up (a circular reference). The reason is shown; rephrase with a different target cell. |
| **Panel is blank** | Close the panel and open it again from **Home → Excel Voice**. |
| **"Error installing add-ins" in Excel's status bar** | Seen during development without any effect on the panel. Ignore it if the panel works. |
| **Something else** | Open **Diagnostics** at the bottom of the panel; it has the technical details of the last steps. Include them if you report a problem. |

---

## Good to know

- **Privacy.** Your recording is sent to Groq for transcription, and the text
  plus your sheet's used range and row-1 headers are sent to Groq to work out
  the command. Cell contents are not sent. Each recording is also saved on
  your computer in `proxy/recordings/` so problems can be replayed; delete
  that folder whenever you like.
- **Undo is one step**, and only for Excel Voice's own last change. If you
  edit those cells by hand afterwards, Undo will overwrite your edits.
- **Undo covers** values, formulas, number formats, font, fill, borders and
  alignment. It does not cover conditional formatting, comments, merged cells
  or data validation.
- **Currency** is formatted with `$`. To change the symbol, edit
  `NUMBER_FORMATS` in `addin/src/taskpane/apply.ts`.
- **Keyboard shortcut.** A shortcut (Ctrl+Shift+Alt+V) is declared but did not
  register on the development machine. Use the Speak button.

---

## How it works

```
Speak in the panel ──audio──▶ proxy ──▶ Groq Whisper ──▶ text
text + sheet layout ────────▶ proxy ──▶ Groq LLM ──▶ command list (JSON)
command list ──▶ checked against a strict schema and safety rules
             ──▶ preview + cells selected ──▶ Apply ──▶ Office.js edits the sheet
```

The language model only *interprets*: it turns your sentence into one of a few
fixed command shapes. It never touches the spreadsheet. The add-in validates
every command (Zod schema plus checks such as "is the target inside its own
range"), shows it to you, and only then runs it with Office.js.

```
addin/   Excel task-pane add-in (TypeScript, Office.js)
proxy/   local Node server holding the Groq key; command schema, parsing, eval
```

## For developers

```
cd proxy
npm test                    # Vitest: schema, address maths, parsing (no network)
npm run eval -- --repeat 3  # replays proxy/eval/cases.json against Groq, 3 runs
```

The evaluation set has 38 spoken and typed commands with expected results.
Current model `openai/gpt-oss-120b` scores 97–100% per run.

- [PROJECT_LOG.md](PROJECT_LOG.md): every significant defect, what it first
  looked like, and the measurement that settled it. Read it first.
- [BRIEF.md](BRIEF.md): the original goals, scope and plan.

## License

MIT: free to use, change and share, with no warranty. See [LICENSE](LICENSE).
