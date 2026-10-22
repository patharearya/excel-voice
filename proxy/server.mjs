// Proxy for the task pane. Holds the Groq key so it never ships in the add-in.
//   POST /transcribe  audio body            -> { text }           (Groq Whisper)
//   POST /interpret   { transcript, context } -> validated commands or a reason (Groq LLM)
// Zero dependencies beyond zod: Node 22+ has http, fetch, FormData, Blob and runs .ts directly.
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { interpret } from "./interpret.ts";

const KEY = process.env.GROQ_API_KEY;
if (!KEY) {
  console.error("GROQ_API_KEY is not set. Get a free key at https://console.groq.com/keys");
  process.exit(1);
}
const PORT = Number(process.env.PORT ?? 8787);
const MODEL = "whisper-large-v3-turbo";

// Which web pages may call this proxy. Unset (local development) = any.
// Hosted: ALLOWED_ORIGINS=https://patharearya.github.io, so only the published
// task pane can use the key from a browser. ponytail: CORS stops other web
// pages, not scripts; add a per-IP rate limit if the free Groq quota gets abused.
const ALLOWED = (process.env.ALLOWED_ORIGINS ?? "*").split(",").map((s) => s.trim());
// Recordings are kept for replay when developing; set SAVE_RECORDINGS=0 when hosted.
const SAVE = process.env.SAVE_RECORDINGS !== "0";
const MAX_BODY = 10 * 1024 * 1024; // 10 MB: minutes of Opus audio, far more than a command

const corsFor = (origin = "") => ({
  "Access-Control-Allow-Origin": ALLOWED.includes("*") ? "*" : ALLOWED.includes(origin) ? origin : ALLOWED[0],
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
  Vary: "Origin",
});

http
  .createServer(async (req, res) => {
    const cors = corsFor(req.headers.origin);
    const json = (res, status, body) =>
      res.writeHead(status, { ...cors, "Content-Type": "application/json" }).end(JSON.stringify(body));
    if (req.method === "OPTIONS") return res.writeHead(204, cors).end();
    // The task pane calls this on open, so a sleeping free-tier host starts waking before the first command.
    if (req.method === "GET" && req.url === "/health") return json(res, 200, { ok: true });
    if (req.method !== "POST") return res.writeHead(404, cors).end();

    const chunks = [];
    let size = 0;
    for await (const c of req) {
      size += c.length;
      if (size > MAX_BODY) return json(res, 413, { ok: false, reason: "request too large" });
      chunks.push(c);
    }
    const body = Buffer.concat(chunks);
    const stamp = new Date().toISOString();

    if (req.url === "/interpret") {
      let input;
      try {
        input = JSON.parse(body.toString("utf8"));
      } catch {
        return json(res, 400, { ok: false, reason: "body is not JSON" });
      }
      let r;
      try {
        r = await interpret(input.transcript ?? "", input.context, { model: input.model });
      } catch (e) {
        // A bad context (missing usedRange) throws in the prompt builder; answer, don't crash the proxy.
        console.error(`[${stamp}] interpret threw:`, e);
        return json(res, 400, { ok: false, reason: `bad request: ${e?.message ?? e}` });
      }
      console.log(`[${stamp}] interpret "${input.transcript}" -> ${r.ok ? JSON.stringify(r.commands) : "REJECTED " + r.reason} (${r.ms} ms)`);
      return json(res, 200, r);
    }

    if (req.url !== "/transcribe") return res.writeHead(404, cors).end();
    const type = req.headers["content-type"] ?? "audio/webm";
    console.log(`[${stamp}] ${body.length} bytes, ${type}`);

    const form = new FormData();
    const ext = ["webm", "ogg", "mp4", "wav", "mp3"].find((x) => type.includes(x)) ?? "webm";
    // Keep every recording so parsing/transcription can be replayed without re-recording.
    if (SAVE) {
      const dir = path.join(import.meta.dirname, "recordings");
      fs.mkdirSync(dir, { recursive: true }); // absent on a fresh clone
      const saved = path.join(dir, stamp.replace(/[:.]/g, "-") + "." + ext);
      fs.writeFileSync(saved, body);
      console.log("  saved " + saved);
    }
    form.append("file", new Blob([body], { type }), `audio.${ext}`);
    form.append("model", MODEL);
    form.append("response_format", "json");

    try {
      const t0 = Date.now();
      const r = await fetch("https://api.groq.com/openai/v1/audio/transcriptions", {
        method: "POST",
        headers: { Authorization: `Bearer ${KEY}` },
        body: form,
      });
      const text = await r.text();
      console.log(`  groq ${r.status} in ${Date.now() - t0} ms: ${text.slice(0, 200)}`);
      res.writeHead(r.status, { ...cors, "Content-Type": "application/json" }).end(text);
    } catch (e) {
      console.error("  groq call failed:", e);
      json(res, 502, { error: String(e) });
    }
  })
  .listen(PORT, () =>
    console.log(`proxy listening on http://localhost:${PORT} (/transcribe via ${MODEL}, /interpret); origins ${ALLOWED.join(",")}; recordings ${SAVE ? "saved" : "off"}`)
  );
