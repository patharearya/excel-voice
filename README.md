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
> computer: you start two small programs from a terminal, then use the panel
> inside Excel. There is no one-click installer yet.

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

- **Windows** with **desktop Excel** (Microsoft 365 or Office 2021 or later).
  Excel on the web and Mac are not supported.
- **A microphone.** Typing commands into the panel also works.
- **Node.js 24 LTS** (22.18 or newer also works). Download the Windows
  installer from <https://nodejs.org> and accept the defaults.
- **Git**, to download this project: <https://git-scm.com/download/win>.
  Alternatively, use the green **Code** button on GitHub → **Download ZIP**,
  and unzip it.
- **A free Groq account** for the API key: <https://console.groq.com>. Sign up
  with Google or email; no card is asked for.

---

## Install (once)

Open a terminal: press the Windows key, type **Terminal** (or **PowerShell**),
press Enter. Run each block below in order.

### 1. Download the project

```
git clone https://github.com/patharearya/excel-voice.git
cd excel-voice
```

### 2. Get a Groq API key

1. Go to <https://console.groq.com/keys> and sign in.
2. Click **Create API Key**, give it any name, and copy the key (it starts
   with `gsk_`). You only see it once.

### 3. Set up the voice service (the "proxy")

The proxy is a small program on your own computer. It holds your Groq key so
the key never sits inside Excel.

```
cd proxy
npm install
copy .env.example .env
notepad .env
```

Notepad opens with one line: `GROQ_API_KEY=gsk_...`. Replace `gsk_...` with
your key, save, and close Notepad.

> Keep the key private. `.env` is excluded from Git, so it won't be uploaded
> if you push changes. If it ever leaks, delete it at console.groq.com and
> make a new one.

### 4. Set up the Excel add-in

```
cd ..\addin
npm install
```

This takes a few minutes the first time.

---

## Start it (every time you want to use it)

You need **two terminal windows**: one for the proxy, one for the add-in.

**Terminal 1: the proxy**

```
cd excel-voice\proxy
npm start
```

Wait for `proxy listening on http://localhost:8787`. Leave this window open.

**Terminal 2: the add-in**

```
cd excel-voice\addin
npm start
```

- **The first time only**, Windows asks to install a "development
  certificate" for `localhost`. Click **Yes** on each prompt. The add-in runs
  over HTTPS on your own machine and needs this to be trusted.
- Excel then opens with the add-in loaded. On the **Home** tab, at the right,
  click **Excel Voice** (in the **Voice** group). The panel opens on the right.

If the panel is blank, the add-in's local web server was still starting. Close
the panel with the **X** and click **Excel Voice** again.

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

**First recording:** Windows or Excel may ask for microphone permission.
Allow it.

---

## Stop it

- In the add-in terminal (Terminal 2), or any terminal in the `addin` folder:
  `npm run stop`. This unloads the add-in and stops its web server.
- In the proxy terminal (Terminal 1), press **Ctrl+C**.

You can then close Excel normally.

---

## Troubleshooting

| What you see | What to do |
|---|---|
| **"No sound was picked up"** | The mic recorded silence. Open **Microphone** at the bottom of the panel and pick another device. Bluetooth headsets (AirPods, etc.) often record silence until they're actively in use. |
| **"Microphone access is blocked"** | Windows Settings → Privacy & security → Microphone: turn on microphone access, and access for desktop apps. |
| **"Can't reach the voice service"** | The proxy isn't running. Start it (Terminal 1 above). |
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
