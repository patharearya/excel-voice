// Replay the eval cases against a model N times. Reports accuracy by command
// type and run-to-run variance, prints every failure with a reason, and saves
// the full record under eval/results/ so later prompt changes compare against it.
//
//   node --env-file=.env eval.ts [--model openai/gpt-oss-120b] [--repeat 3] [--only copy]
import fs from "node:fs";
import path from "node:path";
import { interpret, DEFAULT_MODEL, type SheetContext } from "./interpret.ts";
import type { Command } from "./commands.ts";

type Case = { id: string; type: string; transcript: string; expected: Command[]; source?: string };
type Suite = { context: SheetContext; cases: Case[] };

const args = process.argv.slice(2);
const flag = (name: string, dflt: string) => {
  const i = args.indexOf(`--${name}`);
  return i >= 0 ? args[i + 1] : dflt;
};
const model = flag("model", DEFAULT_MODEL);
const repeat = Number(flag("repeat", "1"));
const only = flag("only", "");
// Free tier is ~30 requests/min; a 2 s gap keeps us under it instead of
// hammering 429s and inflating every latency with retry waits.
const gap = Number(flag("gap", "2000"));

const dir = path.join(import.meta.dirname, "eval");
const suite: Suite = JSON.parse(fs.readFileSync(path.join(dir, "cases.json"), "utf8"));
const cases = only ? suite.cases.filter((c) => c.type === only) : suite.cases;

// Drop null fields and sort keys so {bold:true, italic:null} equals {bold:true}.
const norm = (cmds: Command[]) =>
  JSON.stringify(
    cmds.map((c) =>
      Object.fromEntries(
        Object.entries(c)
          .filter(([, v]) => v !== null)
          .sort(([a], [b]) => a.localeCompare(b))
      )
    )
  );

type Row = { run: number; id: string; type: string; pass: boolean; reason: string; ms: number; got: unknown };
const rows: Row[] = [];

for (let run = 1; run <= repeat; run++) {
  for (const c of cases) {
    const r = await interpret(c.transcript, suite.context, { model });
    await new Promise((res) => setTimeout(res, gap));
    let pass = false;
    let reason = "";
    if (!r.ok) reason = `rejected: ${r.reason}`;
    else if (norm(r.commands) !== norm(c.expected)) reason = `mismatch: got ${norm(r.commands)} expected ${norm(c.expected)}`;
    else pass = true;
    rows.push({ run, id: c.id, type: c.type, pass, reason, ms: r.ms, got: r.ok ? r.commands : r.raw });
    if (!pass) console.log(`FAIL run${run} ${c.id} [${c.type}] "${c.transcript}"\n     ${reason}`);
  }
}

// Report: accuracy per type per run, so variance is visible.
const types = [...new Set(cases.map((c) => c.type))];
const pct = (n: number, d: number) => (d ? `${Math.round((100 * n) / d)}%` : "-");
console.log(`\nmodel ${model}, ${cases.length} cases x ${repeat} run(s)`);
console.log(["type", ...Array.from({ length: repeat }, (_, i) => `run${i + 1}`), "all"].join("\t"));
for (const t of [...types, "total"]) {
  const sel = (r: Row) => t === "total" || r.type === t;
  const perRun = Array.from({ length: repeat }, (_, i) => {
    const rs = rows.filter((r) => r.run === i + 1 && sel(r));
    return pct(rs.filter((r) => r.pass).length, rs.length);
  });
  const all = rows.filter(sel);
  console.log([t, ...perRun, pct(all.filter((r) => r.pass).length, all.length)].join("\t"));
}
const flaky = cases.filter((c) => {
  const rs = rows.filter((r) => r.id === c.id);
  return rs.some((r) => r.pass) && rs.some((r) => !r.pass);
});
if (flaky.length) console.log(`flaky (pass in some runs, fail in others): ${flaky.map((c) => c.id).join(", ")}`);
const msAll = rows.map((r) => r.ms).sort((a, b) => a - b);
console.log(`latency ms: median ${msAll[Math.floor(msAll.length / 2)]}, max ${msAll[msAll.length - 1]}`);

const out = path.join(dir, "results", `${new Date().toISOString().replace(/[:.]/g, "-")}_${model.replace(/[^a-z0-9]/gi, "-")}.json`);
fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, JSON.stringify({ model, repeat, context: suite.context, rows }, null, 2));
console.log(`saved ${out}`);
