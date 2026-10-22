# Excel Voice — brag plan

**What it is:** a free side panel for desktop Excel on Windows: say what you want done to the sheet, check the preview, press Enter.
**Who for:** anyone who spends ages selecting cells and clicking through formatting.
**What sets it apart:** nothing changes until you approve it; one-step Undo for changes Excel's own Ctrl+Z can't undo; free (Groq free tier); one-file install.
**Best claim:** "Talk to Excel." / "Excel's own Ctrl+Z can't undo changes made by an add-in" — so it ships its own Undo.
**Visual hook:** "Stop clicking." → hard cut → "Talk to Excel." with the mic button pulsing.
**Real UI shown:** the actual task pane (real `taskpane.html` structure + real `taskpane.css`, Segoe UI), with the app's real status strings and preview wording from `proxy/plan.ts` `describe()`.
**Tone:** `default` — punchy, playful, clean.
**Share caption:** "Talk to Excel. Say it, check the preview, press Enter."

**Fictional data:** the spreadsheet (office supplies: Item / Qty / Price / Total) is invented. No real user data, keys or hosts appear. The only URL shown is the public install page from the README.

## Visual identity
- Excel green `#107c41` (accent), panel silver-blue `#dae6f7`, ink `#1a2638`, muted `#475a75`, line `#b4c6e0`.
- Segoe UI throughout (the panel's own font). App icon from `addin/assets/icon.svg`.

## Storyboard (21.0s, 30fps, 1920×1080)

| # | Time | Scene | On-screen caption |
|---|---|---|---|
| 1 | 0.0–1.4 | Hook. Big type on blue, a cursor frantically clicking ribbon-like chips behind it. | **Stop clicking.** |
| 1b | 1.4–2.8 | Hard cut, Excel-green field, mic icon pulses. | **Talk to Excel.** |
| 2 | 2.8–6.4 | Excel window slides up: grid + Excel Voice panel. Speak clicked → red Stop, level meter, "Listening…". Transcript types in: "make column C currency and bold the header row". "Working out what to change…" | **Just say it.** |
| 3 | 6.4–9.4 | Preview card: "Format C2:C9: currency", "Format A1:D1: bold". Those cells selected in the sheet. "Check the selected cells, then apply." | **Nothing changes until you approve it.** |
| 4 | 9.4–11.4 | Enter. Prices become $, header goes bold. "Done." Last change bar appears. | **Press Enter. Done.** |
| 5 | 11.4–13.6 | Typed: "Total column D at the bottom" → preview "D10 =SUM(D2:D9)" → applied; formula bar shows the live formula. | **A real formula, not a pasted number.** |
| 6 | 13.6–15.6 | Spoken: "Highlight prices over 100 in yellow" → only the matching cells go yellow. | **Rules, not just ranges.** |
| 7 | 15.6–17.6 | Undo clicked → yellow goes away. "Undone. The cells are back to how they were." | **One-step Undo. Because Ctrl+Z can't.** |
| 8 | 17.6–21.0 | Outro: icon, "Excel Voice", "Free · Desktop Excel on Windows · One-file install", install URL. | |

## Sound
120 bpm, C major (C–G–Am–F). Hook: sparse pad + rising swoosh into the cut. From 2.8s: soft kick, offbeat hats, pluck arpeggio. SFX in key and under the music: mic click, soft typing ticks, preview pop (E5/G5), apply chime (C6/G6), undo reverse blip. Outro: final C chord rings out.
