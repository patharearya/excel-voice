# Excel Voice

Voice commands for desktop Excel on Windows. Free stack: Office.js task pane,
Groq Whisper for speech to text, Groq LLM for command parsing, Node proxy.

Brief: [BRIEF.md](BRIEF.md). Log: [PROJECT_LOG.md](PROJECT_LOG.md), read it first.

## Layout

- `addin/` Office task pane add-in (TypeScript, Yeoman scaffold, XML manifest)
- `proxy/` zero-dependency Node server holding the Groq key

## Step 0 spike: run it

1. Get a free Groq key (no card): https://console.groq.com/keys
2. `proxy/.env`: copy `.env.example`, paste the key.
3. Terminal 1:
   ```
   cd proxy && npm start
   ```
   Expect `proxy listening on http://localhost:8787/transcribe`.
4. Terminal 2 (first run installs a localhost HTTPS cert; accept the prompts):
   ```
   cd addin && npm start
   ```
   Excel opens with the add-in sideloaded. Home tab, "Excel Voice" button opens the pane.
5. In the pane: click **Record**, say something, click **Stop**.
   The status line reports each stage. The transcript should land in A1.
6. Shortcut check: with focus on the sheet, press **Ctrl+Shift+Alt+V**. It should toggle recording.
7. Fill the table in PROJECT_LOG.md with what actually happened. Copy the status text verbatim.

Stop with `npm run stop` in `addin/` (removes the sideload).
