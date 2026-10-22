// Music + SFX as one piece: 120 bpm, C major, beat grid locked to the video cuts (1.4 + 0.5k).
const fs = require("fs");
const SR = 44100, DUR = 21.0, N = Math.round(SR * DUR);
const L = new Float32Array(N), R = new Float32Array(N), VL = new Float32Array(N), VR = new Float32Array(N); // dry, verb send
const mtof = m => 440 * Math.pow(2, (m - 69) / 12);
let seed = 7; const rnd = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff) * 2 - 1;

function add(t0, len, fn, gain, pan = 0, send = 0) {
  const s0 = Math.round(t0 * SR), n = Math.round(len * SR), gl = gain * Math.cos((pan + 1) * Math.PI / 4) * 1.414, gr = gain * Math.sin((pan + 1) * Math.PI / 4) * 1.414;
  for (let i = 0; i < n; i++) { const j = s0 + i; if (j < 0 || j >= N) continue; const v = fn(i / SR); L[j] += v * gl; R[j] += v * gr; VL[j] += v * gl * send; VR[j] += v * gr * send; }
}
const env = (t, a, d) => (t < a ? t / a : Math.exp(-(t - a) / d));

// ---- harmony ----
const BEAT = 0.5, T0 = 1.4;
const CH = { C: [48, 52, 55, 60, 64, 67], G: [43, 50, 55, 59, 62, 67], Am: [45, 52, 57, 60, 64, 69], F: [41, 48, 53, 57, 60, 65] };
const prog = ["C", "G", "Am", "F"];
const chordAt = t => (t < T0 ? "F" : t >= 17.9 ? "C" : prog[Math.floor((t - T0) / 2) % 4]);

// pad: one voice per chord segment, soft attack, gentle detune
function pad(t0, t1, name, g) {
  const notes = CH[name].slice(1);
  add(t0, t1 - t0 + 1.2, t => {
    const e = Math.min(1, t / 0.25) * (t > t1 - t0 ? Math.exp(-(t - (t1 - t0)) / 0.35) : 1);
    let v = 0; for (const m of notes) { const f = mtof(m); v += Math.sin(2 * Math.PI * f * t) + 0.5 * Math.sin(2 * Math.PI * f * 1.003 * t + 1) + 0.15 * Math.sin(4 * Math.PI * f * t); }
    return v * e / notes.length;
  }, g, 0, 0.5);
}
pad(0, 1.4, "F", 0.05);
for (let b = T0; b < 17.9; b += 2) pad(b, Math.min(b + 2, 17.9), chordAt(b + 0.01), 0.07);
pad(17.9, 20.3, "C", 0.09);

// bass: root on beats from 2.9
for (let t = 2.9; t < 17.85; t += BEAT) { const f = mtof(CH[chordAt(t)][0] - 12 + 12); add(t, 0.45, x => (Math.sin(2 * Math.PI * f * x) + 0.3 * Math.sin(4 * Math.PI * f * x)) * env(x, 0.005, 0.18), 0.16); }
add(17.9, 3, x => Math.sin(2 * Math.PI * mtof(36) * x) * env(x, 0.005, 0.9), 0.2);

// pluck arp: eighths, chord tones up an octave
for (let k = 0, t = 2.9; t < 20.5; t += BEAT / 2, k++) {
  const ch = CH[chordAt(t)], m = ch[[3, 4, 5, 4][k % 4]] + 12, f = mtof(m), fade = t > 19 ? Math.max(0, 1 - (t - 19) / 1.5) : 1;
  add(t, 0.5, x => (Math.sin(2 * Math.PI * f * x) + 0.25 * Math.sin(6 * Math.PI * f * x)) * env(x, 0.003, 0.09), 0.05 * fade, k % 2 ? 0.35 : -0.35, 0.35);
}

// drums
const kick = (t, g) => add(t, 0.35, x => Math.sin(2 * Math.PI * (45 * x + 75 * 0.04 * (1 - Math.exp(-x / 0.04)))) * env(x, 0.002, 0.11), g);
const hat = (t, g) => { let lp = 0; add(t, 0.06, x => { const n = rnd(); lp += 0.6 * (n - lp); return (n - lp) * env(x, 0.001, 0.015); }, g, 0.3, 0.1); };
for (let t = 2.9; t < 17.85; t += BEAT) kick(t, 0.3);
for (let t = 3.15; t < 17.85; t += BEAT) hat(t, 0.05);
kick(1.4, 0.32); kick(17.9, 0.4);

// ---- SFX, in key, under the music ----
const click = (t, g = 0.09) => { let lp = 0; add(t, 0.04, x => { const n = rnd(); lp += 0.25 * (n - lp); return (lp * 2 + 0.3 * Math.sin(2 * Math.PI * mtof(96) * x)) * env(x, 0.0005, 0.008); }, g, 0.1, 0.25); };
const blip = (t, m0, m1, g) => { let ph = 0; add(t, 0.25, x => { const f = mtof(m0 + (m1 - m0) * Math.min(1, x / 0.06)); ph += 2 * Math.PI * f / SR; return Math.sin(ph) * env(x, 0.004, 0.06); }, g, 0, 0.4); };
const chime = (t, g) => add(t, 1.2, x => (Math.sin(2 * Math.PI * mtof(84) * x) + 0.6 * Math.sin(2 * Math.PI * mtof(91) * x) * env(x, 0.03, 0.2) + 0.2 * Math.sin(2 * Math.PI * mtof(84) * 2.76 * x) * Math.exp(-x / 0.05)) * env(x, 0.003, 0.35), g, 0, 0.5);
const whoosh = (t, len, g) => { let lp = 0; add(t, len, x => { const k = x / len, n = rnd(); lp += (0.02 + 0.25 * k) * (n - lp); return lp * Math.sin(Math.PI * k) ** 2 * 3; }, g, 0, 0.6); };

for (let t = 0; t < 1.36; t += 0.13) click(t, 0.05);          // the clicking we're replacing
whoosh(0.95, 0.5, 0.12);                                       // riser into the cut
whoosh(2.6, 0.6, 0.08);                                        // window slides in
[3.4, 4.85, 13.95, 16.0, 9.25, 13.0, 15.1].forEach(t => click(t));
[6.4, 12.7, 14.8].forEach(t => blip(t, 76, 79, 0.06));          // preview appears: E5 -> G5
[9.7, 13.15, 15.25].forEach(t => chime(t, 0.06));              // applied: C6 + G6
blip(16.0, 79, 72, 0.06);                                      // undo: G5 -> C5
for (let i = 0; i < 28; i++) click(11.5 + i * 0.8 / 28 + rnd() * 0.005, 0.025); // typing
whoosh(17.6, 0.5, 0.08);
chime(18.1, 0.05);

// ---- shared room: a few combs per channel, then lowpass ----
function verb(src, delays) {
  const out = new Float32Array(N);
  for (const d of delays) { const n = Math.round(d * SR), buf = new Float32Array(N); for (let i = 0; i < N; i++) { buf[i] = src[i] + (i >= n ? buf[i - n] * 0.72 : 0); out[i] += buf[i] / delays.length; } }
  let lp = 0; for (let i = 0; i < N; i++) { lp += 0.25 * (out[i] - lp); out[i] = lp; }
  return out;
}
const rl = verb(VL, [0.0297, 0.0371, 0.0411, 0.0437]), rr = verb(VR, [0.0313, 0.0353, 0.0429, 0.0451]);

// master: sum, soft clip, fade, normalise to -1 dBFS
const buf = Buffer.alloc(44 + N * 4); let peak = 0;
const mix = new Float32Array(N * 2);
for (let i = 0; i < N; i++) {
  const t = i / SR, fade = Math.min(1, t / 0.02) * (t > 20.2 ? Math.max(0, (DUR - t) / 0.8) : 1);
  for (const [c, v] of [[0, L[i] + rl[i] * 0.5], [1, R[i] + rr[i] * 0.5]]) { const y = Math.tanh(v * 1.2) * fade; mix[i * 2 + c] = y; peak = Math.max(peak, Math.abs(y)); }
}
const g = 0.89 / peak;
buf.write("RIFF", 0); buf.writeUInt32LE(36 + N * 4, 4); buf.write("WAVEfmt ", 8); buf.writeUInt32LE(16, 16); buf.writeUInt16LE(1, 20); buf.writeUInt16LE(2, 22);
buf.writeUInt32LE(SR, 24); buf.writeUInt32LE(SR * 4, 28); buf.writeUInt16LE(4, 32); buf.writeUInt16LE(16, 34); buf.write("data", 36); buf.writeUInt32LE(N * 4, 40);
for (let i = 0; i < N * 2; i++) buf.writeInt16LE(Math.round(mix[i] * g * 32767), 44 + i * 2);
fs.writeFileSync(__dirname + "/audio.wav", buf);
console.log("peak before norm", peak.toFixed(3), "gain", g.toFixed(2));
