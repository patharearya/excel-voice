/* global console, document, navigator, fetch, Blob, MediaRecorder, MediaStream, Response, AudioContext, HTMLSelectElement, HTMLInputElement, HTMLButtonElement, requestAnimationFrame, Office */

// Task pane: mic (or typed text) -> proxy /transcribe -> proxy /interpret ->
// preview (cells selected, nothing changed) -> Apply / Cancel -> Undo.
import type { Command } from "../../../proxy/commands.ts";
import { describe, problem } from "../../../proxy/plan.ts";
import { sheetContext, select, apply, undo, type Undo } from "./apply";

const PROXY = "http://localhost:8787";

let pending: Command[] = []; // previewed, not yet applied
let last: Undo | null = null; // one-step Undo

let rec: MediaRecorder | null = null;
let chunks: Blob[] = [];
let startedAt = 0;
let peak = 0; // loudest sample seen this recording, 0..1
let micInfo = "";
let inExcel = false;

const $ = (id: string) => document.getElementById(id)!;
const log = (msg: string) => {
  $("status").textContent = msg;
  console.log(msg);
};

Office.onReady((info) => {
  // Also runs in a plain browser (host null) as a control for the Excel run.
  inExcel = info.host === Office.HostType.Excel;
  $("sideload-msg").style.display = "none";
  $("app-body").style.display = "block";
  $("record").onclick = () => toggle();
  $("typed").onkeydown = (e) => {
    if (e.key === "Enter") understand(($("typed") as HTMLInputElement).value);
  };
  $("apply").onclick = () => doApply();
  $("cancel").onclick = () => showPreview([], "Cancelled.");
  $("undo").onclick = () => doUndo();
  // Enter applies, Escape cancels, while a preview is showing and focus is not in the text box.
  document.onkeydown = (e) => {
    if (!pending.length || e.target === $("typed")) return;
    if (e.key === "Enter") doApply();
    if (e.key === "Escape") showPreview([], "Cancelled.");
  };
  listMics();
  shortcutDiag();
  // Evidence for the spike questions, visible without devtools.
  log(
    `ready (${inExcel ? "in Excel" : "not in Excel"})\nMediaRecorder: ${typeof MediaRecorder}\nmediaDevices: ${!!navigator.mediaDevices}\n` +
      `webm supported: ${typeof MediaRecorder !== "undefined" && MediaRecorder.isTypeSupported("audio/webm")}\n` +
      `UA: ${navigator.userAgent}`
  );
});

// Is the shared runtime up and did Excel register our shortcut?
async function shortcutDiag() {
  const req = Office.context.requirements;
  let shortcuts = "n/a";
  try {
    shortcuts = JSON.stringify(await Office.actions.getShortcuts());
  } catch (e: any) {
    shortcuts = "getShortcuts failed: " + (e?.message || e);
  }
  $("diag").textContent =
    "SharedRuntime 1.1: " +
    req.isSetSupported("SharedRuntime", "1.1") +
    " | KeyboardShortcuts 1.1: " +
    req.isSetSupported("KeyboardShortcuts", "1.1") +
    " | Office.addin: " +
    typeof Office.addin +
    " | registered shortcuts: " +
    shortcuts;
}

async function listMics() {
  const sel = $("mic") as HTMLSelectElement;
  const keep = sel.value;
  const devs = (await navigator.mediaDevices.enumerateDevices()).filter(
    (d) => d.kind === "audioinput"
  );
  sel.innerHTML =
    '<option value="">(default)</option>' +
    devs
      .map(
        (d) =>
          `<option value="${d.deviceId}">${d.label || "(label hidden until permission granted)"}</option>`
      )
      .join("");
  sel.value = keep;
  $("mics").textContent =
    `${devs.length} audio input(s): ${devs.map((d) => d.label || d.deviceId.slice(0, 8)).join(" | ")}`;
}

async function toggle(source = "button") {
  if (rec) {
    rec.stop();
    return;
  }
  try {
    const deviceId = ($("mic") as HTMLSelectElement).value;
    const stream = await navigator.mediaDevices.getUserMedia({
      audio: deviceId ? { deviceId: { exact: deviceId } } : true,
    });
    listMics(); // labels become visible once permission is granted
    const track = stream.getAudioTracks()[0];
    micInfo =
      `mic: "${track.label}" enabled=${track.enabled} muted=${track.muted} readyState=${track.readyState}\n` +
      `settings: ${JSON.stringify(track.getSettings())}`;
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
      $("record").textContent = "Record";
      const secs = ((Date.now() - startedAt) / 1000).toFixed(1);
      await transcribe(blob, `${secs}s, peak level ${(peak * 100).toFixed(0)}%`);
    };
    rec.start();
    meter(stream);
    $("record").textContent = "Stop";
    log(`recording (started by ${source})… click Stop or press the shortcut again\n${micInfo}`);
  } catch (e: any) {
    // Question 1 is answered here if it fails: name + message is the evidence.
    log(`mic failed: ${e?.name}: ${e?.message}`);
    rec = null;
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
    $("level").style.width = `${Math.round(max * 100)}%`;
    requestAnimationFrame(tick);
  };
  tick();
}

async function transcribe(blob: Blob, info: string) {
  log(`sending ${blob.size} bytes (${blob.type}), ${info}…`);
  let r: Response;
  try {
    r = await fetch(`${PROXY}/transcribe`, { method: "POST", headers: { "Content-Type": blob.type }, body: blob });
  } catch (e: any) {
    log(`proxy unreachable: ${e?.message}. Is "npm start" running in proxy/?`);
    return;
  }
  const body = await r.json();
  if (!r.ok) {
    log(`proxy ${r.status}: ${JSON.stringify(body)}`);
    return;
  }
  log(`${blob.size} bytes, ${info}\n${micInfo}`);
  await understand(body.text);
}

// Transcript -> validated commands -> preview. Never guesses: anything the
// proxy rejects, or that plan.ts flags, is shown with its reason and no Apply.
async function understand(text: string) {
  text = text.trim();
  $("transcript").textContent = text;
  if (!text) return;
  if (!inExcel) return showPreview([], "Not in Excel: nothing to interpret against.");
  let r: any;
  try {
    const context = await sheetContext();
    showPreview([], "Interpreting…");
    const res = await fetch(`${PROXY}/interpret`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ transcript: text, context }),
    });
    r = await res.json();
  } catch (e: any) {
    return showPreview([], `Could not interpret: ${e?.message}. Is "npm start" running in proxy/?`);
  }
  console.log("interpret", text, r);
  if (!r.ok) return showPreview([], `I didn't understand that (${r.reason}).`);
  if (!r.commands.length) return showPreview([], "I didn't understand that as a command.");
  const issues = r.commands.map(problem).filter(Boolean);
  if (issues.length) return showPreview([], `Can't run this: ${issues.join("; ")}`, r.commands);
  showPreview(r.commands, "Apply (Enter) or Cancel (Esc)?");
  try {
    await select(r.commands);
  } catch (e: any) {
    log(`could not select the preview cells: ${e?.message}`);
  }
}

// `cmds` can be applied; `shown` is only listed (e.g. a command with a problem).
function showPreview(cmds: Command[], message: string, shown: Command[] = cmds) {
  pending = cmds;
  $("preview").textContent = [...shown.map(describe), message].join("\n");
  ($("apply") as HTMLButtonElement).disabled = !cmds.length;
  ($("cancel") as HTMLButtonElement).disabled = !cmds.length;
}

async function doApply() {
  const cmds = pending;
  if (!cmds.length) return;
  showPreview([], "Applying…", cmds);
  try {
    const r = await apply(cmds);
    last = r.undo.snaps.length ? r.undo : null;
    $("last").textContent = r.undo.done.length ? `Last: ${r.undo.done.join("; ")}` : "";
    ($("undo") as HTMLButtonElement).disabled = !last;
    showPreview([], r.error ? `Stopped: ${r.error}. Undo reverts what ran.` : "Done.", cmds);
  } catch (e: any) {
    showPreview([], `Apply failed: ${e?.message}`, cmds);
  }
}

async function doUndo() {
  if (!last) return;
  try {
    await undo(last);
    last = null;
    $("last").textContent = "Undone.";
    ($("undo") as HTMLButtonElement).disabled = true;
  } catch (e: any) {
    $("last").textContent = `Undo failed: ${e?.message}`;
  }
}

// Keyboard shortcut target (needs shared runtime + shortcuts.json in the manifest).
Office.actions.associate("toggleRecord", async (event: Office.AddinCommands.Event) => {
  log("shortcut fired");
  await toggle("shortcut");
  event.completed();
});
