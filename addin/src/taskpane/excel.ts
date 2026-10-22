/* global console, document, navigator, fetch, Blob, MediaRecorder, MediaStream, Response, AudioContext, HTMLSelectElement, HTMLInputElement, HTMLButtonElement, MediaDeviceInfo, requestAnimationFrame, Office */

// Task pane: mic (or typed text) -> proxy /transcribe -> proxy /interpret ->
// preview (cells selected, nothing changed) -> Apply / Cancel -> Undo.
import type { Command } from "../../../proxy/commands.ts";
import { describe, problem } from "../../../proxy/plan.ts";
import { sheetContext, select, apply, undo, type Undo } from "./apply";

// Served from localhost (npm start) -> the local proxy; served from GitHub Pages
// -> the hosted proxy on Render. Render's URL is set in render.yaml's service name.
const HOSTED_PROXY = "https://excel-voice-proxy.onrender.com";
const PROXY = location.hostname === "localhost" ? "http://localhost:8787" : HOSTED_PROXY;
// Below this peak level a recording is treated as silence and not sent:
// silent audio comes back from Whisper as "Thank you." (Defect 1).
const SILENT_PEAK = 0.02;

let pending: Command[] = []; // previewed, not yet applied
let last: Undo | null = null; // one-step Undo
let busy = false; // a transcribe/interpret/apply is in flight

let rec: MediaRecorder | null = null;
let chunks: Blob[] = [];
let startedAt = 0;
let peak = 0; // loudest sample seen this recording, 0..1
let inExcel = false;

const $ = (id: string) => document.getElementById(id)!;
const btn = (id: string) => $(id) as HTMLButtonElement;

type Tone = "idle" | "busy" | "ok" | "warn" | "error";
function status(text: string, tone: Tone = "idle") {
  $("status").textContent = text;
  $("status").dataset.tone = tone;
}

// Everything technical goes here, never into the main status line.
const diagLines: string[] = [];
function diag(line: string) {
  console.log(line);
  diagLines.push(`${new Date().toLocaleTimeString()} ${line}`);
  $("diag").textContent = diagLines.slice(-40).join("\n");
}

Office.onReady((info) => {
  // Also runs in a plain browser (host null), for checking the layout.
  inExcel = info.host === Office.HostType.Excel;
  applyTheme();
  $("sideload-msg").hidden = true;
  $("app-body").hidden = false;
  btn("record").onclick = () => toggle();
  $("typed").onkeydown = (e) => {
    const input = $("typed") as HTMLInputElement;
    if (e.key === "Enter" && input.value.trim()) {
      understand(input.value);
      input.value = "";
    }
  };
  document.querySelectorAll<HTMLButtonElement>(".example").forEach((b) => (b.onclick = () => understand(b.textContent!)));
  btn("apply").onclick = () => doApply();
  btn("cancel").onclick = () => cancel();
  btn("undo").onclick = () => doUndo();
  // Enter applies, Escape cancels, while a preview is showing and focus is not in the text box.
  document.onkeydown = (e) => {
    if (!pending.length || e.target === $("typed")) return;
    if (e.key === "Enter") {
      e.preventDefault();
      doApply();
    }
    if (e.key === "Escape") cancel();
  };
  listMics();
  diag(`ready, ${inExcel ? "in Excel" : "not in Excel"}; ${navigator.userAgent}`);
  shortcutDiag();
  status(inExcel ? "Ready" : "Not running inside Excel: commands can't be applied here.", inExcel ? "idle" : "warn");
  if (inExcel) wake();
});

// The free hosted proxy sleeps after 15 idle minutes and takes up to a minute
// to wake. Ping it as soon as the pane opens, and say so if it is slow.
async function wake() {
  const slow = setTimeout(() => status("Waking up the voice service. The first time can take up to a minute", "busy"), 2000);
  try {
    const r = await fetch(`${PROXY}/health`);
    diag(`proxy ${PROXY} health ${r.status}`);
    if ($("status").dataset.tone === "busy" && !busy) status("Ready");
  } catch (e: any) {
    diag(`proxy ${PROXY} unreachable: ${e?.message}`);
    status(UNREACHABLE, "error");
  } finally {
    clearTimeout(slow);
  }
}

// Dark Office themes (dark grey, black) take Office's own colours. Light themes
// keep the CSS paper background on purpose, so the pane doesn't blend into the grid.
function applyTheme() {
  const t = Office.context?.officeTheme;
  if (!t) return;
  const hex = t.bodyBackgroundColor.replace("#", "");
  const lum = (0.299 * parseInt(hex.slice(0, 2), 16) + 0.587 * parseInt(hex.slice(2, 4), 16) + 0.114 * parseInt(hex.slice(4, 6), 16)) / 255;
  if (lum >= 0.5) return;
  document.documentElement.dataset.theme = "dark";
  const root = document.documentElement.style;
  root.setProperty("--bg", t.bodyBackgroundColor);
  root.setProperty("--fg", t.bodyForegroundColor);
  root.setProperty("--surface", t.controlBackgroundColor);
  root.setProperty("--field", t.controlBackgroundColor);
  root.setProperty("--line", t.controlForegroundColor + "40");
}

// Is the shared runtime up and did Excel register our shortcut? (Defect 2, parked.)
async function shortcutDiag() {
  const req = Office.context.requirements;
  let shortcuts = "n/a";
  try {
    shortcuts = JSON.stringify(await Office.actions.getShortcuts());
  } catch (e: any) {
    shortcuts = "getShortcuts failed: " + (e?.message || e);
  }
  diag(
    `SharedRuntime 1.1: ${req.isSetSupported("SharedRuntime", "1.1")}, KeyboardShortcuts 1.1: ${req.isSetSupported("KeyboardShortcuts", "1.1")}, registered shortcuts: ${shortcuts}`
  );
}

async function listMics() {
  const sel = $("mic") as HTMLSelectElement;
  const keep = sel.value;
  let devs: MediaDeviceInfo[];
  try {
    devs = (await navigator.mediaDevices.enumerateDevices()).filter((d) => d.kind === "audioinput");
  } catch (e: any) {
    return diag(`listing microphones failed: ${e?.message}`);
  }
  sel.innerHTML =
    '<option value="">System default</option>' +
    devs.map((d, i) => `<option value="${d.deviceId}">${d.label || `Microphone ${i + 1}`}</option>`).join("");
  sel.value = keep;
}

function recording(on: boolean) {
  const b = btn("record");
  if (on) b.dataset.recording = "";
  else delete b.dataset.recording;
  $("record-label").textContent = on ? "Stop" : "Speak";
  $("meter").hidden = !on;
}

async function toggle() {
  if (rec) {
    rec.stop();
    return;
  }
  if (busy) return;
  try {
    const deviceId = ($("mic") as HTMLSelectElement).value;
    const stream = await navigator.mediaDevices.getUserMedia({
      audio: deviceId ? { deviceId: { exact: deviceId } } : true,
    });
    listMics(); // labels become visible once permission is granted
    const track = stream.getAudioTracks()[0];
    diag(`mic "${track.label}" muted=${track.muted} settings=${JSON.stringify(track.getSettings())}`);
    chunks = [];
    peak = 0;
    startedAt = Date.now();
    const mimeType = MediaRecorder.isTypeSupported("audio/webm") ? "audio/webm" : undefined;
    rec = new MediaRecorder(stream, mimeType ? { mimeType } : undefined);
    rec.ondataavailable = (e) => chunks.push(e.data);
    rec.onstop = async () => {
      stream.getTracks().forEach((t) => t.stop());
      const blob = new Blob(chunks, { type: rec!.mimeType });
      rec = null;
      recording(false);
      const secs = (Date.now() - startedAt) / 1000;
      diag(`recorded ${blob.size} bytes, ${secs.toFixed(1)} s, peak ${(peak * 100).toFixed(0)}%`);
      if (peak < SILENT_PEAK)
        return status("No sound was picked up. Check the microphone under Microphone below, then try again.", "warn");
      await transcribe(blob);
    };
    rec.start();
    meter(stream);
    recording(true);
    status("Listening. Click Stop when you're done", "busy");
  } catch (e: any) {
    rec = null;
    recording(false);
    diag(`mic failed: ${e?.name}: ${e?.message}`);
    if (e?.name === "NotAllowedError")
      status("Microphone access is blocked. Allow it in Windows Settings > Privacy & security > Microphone.", "error");
    else if (e?.name === "NotFoundError" || e?.name === "OverconstrainedError")
      status("That microphone isn't available. Pick another one under Microphone below.", "error");
    else status(`Couldn't start the microphone (${e?.name}).`, "error");
  }
}

// Live input level so a silent mic is visible while recording, not after.
function meter(stream: MediaStream) {
  const ctx = new AudioContext();
  const analyser = ctx.createAnalyser();
  ctx.createMediaStreamSource(stream).connect(analyser);
  const buf = new Uint8Array(analyser.fftSize);
  const tick = () => {
    if (!rec) {
      ctx.close();
      return;
    }
    analyser.getByteTimeDomainData(buf);
    let max = 0;
    for (const v of buf) max = Math.max(max, Math.abs(v - 128) / 128);
    peak = Math.max(peak, max);
    $("level").style.transform = `scaleX(${max.toFixed(2)})`;
    requestAnimationFrame(tick);
  };
  tick();
}

const UNREACHABLE =
  PROXY === HOSTED_PROXY
    ? "Can't reach the voice service. Check your internet connection and try again."
    : "Can't reach the voice service. Is the proxy running (npm start in proxy/)?";
const RATE_LIMITED = "Too many requests right now. Wait a few seconds and try again.";

async function transcribe(blob: Blob) {
  busy = true;
  status("Transcribing", "busy");
  try {
    let r: Response;
    try {
      r = await fetch(`${PROXY}/transcribe`, { method: "POST", headers: { "Content-Type": blob.type }, body: blob });
    } catch (e: any) {
      diag(`transcribe fetch failed: ${e?.message}`);
      return status(UNREACHABLE, "error");
    }
    const body = await r.json().catch(() => ({}));
    if (!r.ok) {
      diag(`transcribe ${r.status}: ${JSON.stringify(body)}`);
      return status(r.status === 429 ? RATE_LIMITED : `Transcription failed (${r.status}).`, "error");
    }
    busy = false;
    await understand(body.text ?? "");
  } finally {
    busy = false;
  }
}

// Transcript -> validated commands -> preview. Never guesses: anything the
// proxy rejects, or that plan.ts flags, is shown with its reason and no Apply.
async function understand(text: string) {
  text = text.trim();
  if (!text || busy) return;
  if (rec) return;
  showTranscript(text);
  clearPreview();
  if (!inExcel) return status("Not running inside Excel: commands can't be applied here.", "warn");
  busy = true;
  status("Working out what to change", "busy");
  try {
    let r: any;
    try {
      const context = await sheetContext();
      const res = await fetch(`${PROXY}/interpret`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ transcript: text, context }),
      });
      r = await res.json();
    } catch (e: any) {
      diag(`interpret failed: ${e?.message}`);
      return status(UNREACHABLE, "error");
    }
    diag(`interpret "${text}" -> ${JSON.stringify(r.ok ? r.commands : r.reason)}`);
    if (!r.ok)
      return status(/\b429\b/.test(r.reason) ? RATE_LIMITED : "Something went wrong reading that. Try saying it again.", "error");
    if (!r.commands.length)
      return status("I didn't catch a command in that. The examples below show what I can do.", "warn");
    const issues: string[] = r.commands.map(problem).filter(Boolean);
    if (issues.length) {
      showPreview(r.commands, true);
      return status(`Can't do this safely: ${issues.join("; ")}.`, "error");
    }
    showPreview(r.commands, false);
    status("Check the selected cells, then apply.", "ok");
    try {
      await select(r.commands);
    } catch (e: any) {
      diag(`select failed: ${e?.message}`);
    }
    btn("apply").focus();
  } finally {
    busy = false;
  }
}

function showTranscript(text: string) {
  $("transcript").textContent = text;
  $("transcript").hidden = false;
}

// `blocked`: list the commands but offer no Apply (a problem was found).
function showPreview(cmds: Command[], blocked: boolean) {
  pending = blocked ? [] : cmds;
  const list = $("changes");
  list.replaceChildren(
    ...cmds.map((c) => {
      const li = document.createElement("li");
      li.textContent = describe(c);
      return li;
    })
  );
  if (blocked) $("preview").dataset.blocked = "";
  else delete $("preview").dataset.blocked;
  $("preview").querySelector(".actions")!.toggleAttribute("hidden", blocked);
  $("preview").hidden = false;
  $("examples").hidden = true;
}

function clearPreview() {
  pending = [];
  $("preview").hidden = true;
}

function cancel() {
  clearPreview();
  $("examples").hidden = false;
  status("Cancelled. Nothing was changed.");
}

async function doApply() {
  const cmds = pending;
  if (!cmds.length || busy) return;
  busy = true;
  clearPreview();
  status("Applying", "busy");
  try {
    const r = await apply(cmds);
    last = r.undo.snaps.length ? r.undo : null;
    $("last").textContent = r.undo.done.join("; ") || "Nothing";
    $("last-bar").hidden = !last;
    btn("undo").disabled = !last;
    if (r.error) {
      diag(`apply stopped: ${r.error}`);
      status(`Stopped partway: ${r.error}. Undo reverts what ran.`, "error");
    } else status("Done.", "ok");
  } catch (e: any) {
    diag(`apply failed: ${e?.message}`);
    status(`Couldn't apply that: ${e?.message}`, "error");
  } finally {
    busy = false;
    $("examples").hidden = false;
  }
}

async function doUndo() {
  if (!last || busy) return;
  busy = true;
  try {
    await undo(last);
    last = null;
    $("last-bar").hidden = true;
    status("Undone. The cells are back to how they were.", "ok");
  } catch (e: any) {
    diag(`undo failed: ${e?.message}`);
    status(`Undo failed: ${e?.message}`, "error");
  } finally {
    busy = false;
  }
}

// Keyboard shortcut target (needs shared runtime + shortcuts.json in the manifest).
Office.actions.associate("toggleRecord", async (event: Office.AddinCommands.Event) => {
  diag("shortcut fired");
  await toggle();
  event.completed();
});
