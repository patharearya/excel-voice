// Pure address arithmetic for a validated command: which cells it changes (what
// the preview selects and Undo snapshots), what it says in the preview, and
// whether it is safe to run. No Excel here, so it is tested with Vitest; the
// add-in imports it directly.
import type { Command } from "./commands.ts";

// Excel's grid limits (XFD1048576). The schema regex allows larger addresses.
const MAX_COL = 16384;
const MAX_ROW = 1048576;

export const colNum = (col: string) => [...col].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
export const colName = (n: number): string =>
  n <= 0 ? "" : colName(Math.floor((n - 1) / 26)) + String.fromCharCode(65 + ((n - 1) % 26));

export type Box = { c1: number; r1: number; c2: number; r2: number };

// "C50:C2" and "C2:C50" give the same box.
export function box(range: string): Box {
  const cell = (s: string) => {
    const m = /^([A-Z]+)([0-9]+)$/.exec(s);
    if (!m) throw new Error(`not a cell: ${s}`);
    return [colNum(m[1]), Number(m[2])];
  };
  const [a, b = a] = range.split(":");
  const [ca, ra] = cell(a);
  const [cb, rb] = cell(b);
  return { c1: Math.min(ca, cb), r1: Math.min(ra, rb), c2: Math.max(ca, cb), r2: Math.max(ra, rb) };
}

export function address({ c1, r1, c2, r2 }: Box): string {
  const a = colName(c1) + r1;
  const b = colName(c2) + r2;
  return a === b ? a : `${a}:${b}`;
}

export function size(range: string) {
  const b = box(range);
  return { rows: b.r2 - b.r1 + 1, cols: b.c2 - b.c1 + 1 };
}

// The cells a command changes.
export function touched(cmd: Command): string {
  switch (cmd.action) {
    case "copy": {
      const f = box(cmd.from);
      const t = box(cmd.to);
      return address({ c1: t.c1, r1: t.r1, c2: t.c1 + f.c2 - f.c1, r2: t.r1 + f.r2 - f.r1 });
    }
    case "format":
    case "sort":
    case "arithmetic":
      return address(box(cmd.range));
    case "compute":
      return cmd.target;
  }
}

const FN = { sum: "SUM", average: "AVERAGE", min: "MIN", max: "MAX", count: "COUNT" } as const;
export const formula = (cmd: Extract<Command, { action: "compute" }>) => `=${FN[cmd.op]}(${address(box(cmd.range))})`;

// One preview line per command, in the user's terms.
export function describe(cmd: Command): string {
  switch (cmd.action) {
    case "copy":
      return `Copy ${address(box(cmd.from))} to ${touched(cmd)}`;
    case "format": {
      const what = [
        cmd.bold !== null && (cmd.bold ? "bold" : "not bold"),
        cmd.italic !== null && (cmd.italic ? "italic" : "not italic"),
        cmd.numberFormat,
        cmd.fontColor && `font ${cmd.fontColor}`,
        cmd.fillColor && `fill ${cmd.fillColor}`,
      ].filter(Boolean);
      return `Format ${touched(cmd)}: ${what.join(", ") || "nothing"}`;
    }
    case "compute":
      return `${cmd.target} ${formula(cmd)}`;
    case "sort":
      return `Sort ${touched(cmd)} by column ${cmd.by}, ${cmd.order === "asc" ? "ascending" : "descending"}${cmd.hasHeader ? ", keeping the header row" : ""}`;
    case "arithmetic":
      return {
        add: `Add ${cmd.value} to ${touched(cmd)}`,
        subtract: `Subtract ${cmd.value} from ${touched(cmd)}`,
        multiply: `Multiply ${touched(cmd)} by ${cmd.value}`,
        divide: `Divide ${touched(cmd)} by ${cmd.value}`,
      }[cmd.op];
  }
}

const SYM = { add: "+", subtract: "-", multiply: "*", divide: "/" } as const;

// New cell contents for an arithmetic command, from the range's current
// `formulas`. Numbers are computed; formulas are wrapped so they stay live
// (=B2*C2 -> =(B2*C2)*2); text, blanks and booleans are left as they are.
export function arithmetic(formulas: unknown[][], cmd: Extract<Command, { action: "arithmetic" }>): unknown[][] {
  const f = (n: number) =>
    cmd.op === "add" ? n + cmd.value : cmd.op === "subtract" ? n - cmd.value : cmd.op === "multiply" ? n * cmd.value : n / cmd.value;
  return formulas.map((row) =>
    row.map((v) => {
      if (typeof v === "number") return f(v);
      if (typeof v === "string" && v.startsWith("=")) return `=(${v.slice(1)})${SYM[cmd.op]}${cmd.value}`;
      return v;
    })
  );
}

// Why a schema-valid command must not run, or null. Shown instead of Apply.
export function problem(cmd: Command): string | null {
  const ranges = [touched(cmd), ...(cmd.action === "copy" ? [cmd.from] : []), ...(cmd.action === "compute" || cmd.action === "sort" ? [cmd.range] : [])];
  for (const r of ranges) {
    const b = box(r);
    if (b.c2 > MAX_COL || b.r2 > MAX_ROW) return `${address(b)} is outside the sheet`;
  }
  if (cmd.action === "format" && [cmd.bold, cmd.italic, cmd.numberFormat, cmd.fontColor, cmd.fillColor].every((v) => v === null))
    return `format ${cmd.range} names no change`;
  if (cmd.action === "compute") {
    const t = box(cmd.target);
    const r = box(cmd.range);
    if (t.c1 >= r.c1 && t.c1 <= r.c2 && t.r1 >= r.r1 && t.r1 <= r.r2) return `${cmd.target} is inside ${cmd.range}: circular reference`;
  }
  if (cmd.action === "arithmetic" && cmd.op === "divide" && cmd.value === 0) return "division by zero";
  if (cmd.action === "sort") {
    const c = colNum(cmd.by);
    const r = box(cmd.range);
    if (c < r.c1 || c > r.c2) return `column ${cmd.by} is not in ${cmd.range}`;
  }
  return null;
}
