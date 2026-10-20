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

const cors = {
  "Access-Control-Allow-Origin": "*", // ponytail: wide open on localhost; pin to the add-in origin before hosting on Render
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};
const json = (res, status, body) =>
  res.writeHead(status, { ...cors, "Content-Type": "application/json" }).end(JSON.stringify(body));

http
  .createServer(async (req, res) => {
    if (req.method === "OPTIONS") return res.writeHead(204, cors).end();
    if (req.method !== "POST") return res.writeHead(404, cors).end();

    const chunks = [];
    for await (const c of req) chunks.push(c);
    const body = Buffer.concat(chunks);
    const stamp = new Date().toISOString();

    if (req.url === "/interpret") {
      let input;
      try {
        input = JSON.parse(body.toString("utf8"));
      } catch {
        return json(res, 400, { ok: false, reason: "body is not JSON" });
      }
      const r = await interpret(input.transcript ?? "", input.context, { model: input.model });
      console.log(`[${stamp}] interpret "${input.transcript}" -> ${r.ok ? JSON.stringify(r.commands) : "REJECTED " + r.reason} (${r.ms} ms)`);
      return json(res, 200, r);
    }

    if (req.url !== "/transcribe") return res.writeHead(404, cors).end();
    const type = req.headers["content-type"] ?? "audio/webm";
    console.log(`[${stamp}] ${body.length} bytes, ${type}`);

    const form = new FormData();
    const ext = ["webm", "ogg", "mp4", "wav", "mp3"].find((x) => type.includes(x)) ?? "webm";
    // Keep every recording so parsing/transcription can be replayed without re-recording.
    const saved = path.join(import.meta.dirname, "recordings", stamp.replace(/[:.]/g, "-") + "." + ext);
    fs.writeFileSync(saved, body);
    console.log("  saved " + saved);
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
  .listen(PORT, () => console.log(`proxy listening on http://localhost:${PORT} (/transcribe via ${MODEL}, /interpret)`));
