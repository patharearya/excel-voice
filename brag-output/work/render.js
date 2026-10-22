// node render.js stills t1 t2 ...  -> stills/t.png
// node render.js video             -> video.mp4 (silent, 30fps)
const http = require("http"), fs = require("fs"), path = require("path"), { spawn } = require("child_process");
const puppeteer = require("puppeteer-core"), ffmpeg = require("ffmpeg-static");
const FPS = 30, DUR = 21.0, DIR = __dirname;
const CHROME = "C:/Program Files/Google/Chrome/Application/chrome.exe";

const srv = http.createServer((q, r) => {
  const f = path.join(DIR, decodeURIComponent(q.url.split("?")[0]));
  fs.readFile(f, (e, d) => { if (e) { r.writeHead(404); return r.end(); } r.writeHead(200, { "Content-Type": { ".html": "text/html", ".css": "text/css", ".svg": "image/svg+xml" }[path.extname(f)] || "application/octet-stream" }); r.end(d); });
}).listen(0);

(async () => {
  const port = srv.address().port;
  const b = await puppeteer.launch({ executablePath: CHROME, headless: true, args: ["--force-device-scale-factor=1", "--hide-scrollbars"] });
  const p = await b.newPage();
  await p.setViewport({ width: 1920, height: 1080 });
  await p.goto(`http://localhost:${port}/video.html`);
  await p.evaluate(() => window.ready);
  const [mode, ...ts] = process.argv.slice(2);
  if (mode === "stills") {
    fs.mkdirSync(path.join(DIR, "stills"), { recursive: true });
    for (const t of ts) { await p.evaluate(t => render(+t), t); await p.screenshot({ path: path.join(DIR, "stills", `${t}.png`) }); }
  } else {
    const ff = spawn(ffmpeg, ["-y", "-f", "image2pipe", "-framerate", String(FPS), "-i", "-", "-c:v", "libx264", "-pix_fmt", "yuv420p", "-crf", "16", "-preset", "medium", path.join(DIR, "video.mp4")], { stdio: ["pipe", "ignore", "inherit"] });
    const n = Math.round(DUR * FPS);
    for (let i = 0; i < n; i++) {
      await p.evaluate(t => render(t), i / FPS);
      const buf = await p.screenshot({ type: "jpeg", quality: 95 });
      if (!ff.stdin.write(buf)) await new Promise(r => ff.stdin.once("drain", r));
    }
    ff.stdin.end();
    await new Promise(r => ff.on("close", r));
  }
  await b.close(); srv.close();
})();
