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
    case "clear":
    case "set":
      return address(box(cmd.range));
    case "compute":
      return cmd.target;
  }
}

const FN = { sum: "SUM", average: "AVERAGE", min: "MIN", max: "MAX", count: "COUNT", counta: "COUNTA", median: "MEDIAN" } as const;

type Format = Extract<Command, { action: "format" }>;

// Excel number-format code for a spoken format. Defaults match what the
// add-in used before decimals existed: currency/number 2 places, percent 0.
// ponytail: US-style currency and date; change here for another locale.
export function numberFormatCode(fmt: Format["numberFormat"], decimals: number | null): string {
  const dec = (d: number) => (d > 0 ? "." + "0".repeat(d) : "");
  switch (fmt ?? "number") {
    case "currency":
      return `$#,##0${dec(decimals ?? 2)}`;
    case "percent":
      return `0${dec(decimals ?? 0)}%`;
    case "date":
      return "m/d/yyyy";
    default:
      return `#,##0${dec(decimals ?? 2)}`;
  }
}

const BORDER = { all: "all borders", outline: "outline border", bottom: "bottom border", none: "no borders" } as const;
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
        cmd.underline !== null && (cmd.underline ? "underline" : "no underline"),
        cmd.fontSize !== null && `size ${cmd.fontSize}`,
        cmd.align && `align ${cmd.align}`,
        cmd.wrap !== null && (cmd.wrap ? "wrap text" : "no wrap"),
        cmd.border && BORDER[cmd.border],
        cmd.numberFormat,
        cmd.decimals !== null && `${cmd.decimals} decimal${cmd.decimals === 1 ? "" : "s"}`,
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
        round: `Round ${touched(cmd)} to ${cmd.value} decimal${cmd.value === 1 ? "" : "s"}`,
      }[cmd.op];
    case "clear":
      return { contents: `Clear the contents of ${touched(cmd)}`, formats: `Remove formatting from ${touched(cmd)}`, all: `Clear ${touched(cmd)} completely` }[cmd.what];
    case "set":
      return `Put ${typeof cmd.value === "string" ? `"${cmd.value}"` : cmd.value} in ${touched(cmd)}`;
  }
}

const SYM = { add: "+", subtract: "-", multiply: "*", divide: "/" } as const;

// Decimal rounding via exponent strings, so 1.005 -> 1.01 (Math.round(100.49999) would give 1).
const roundTo = (n: number, d: number) => Number(Math.round(Number(`${n}e${d}`)) + `e-${d}`);

// New cell contents for an arithmetic command, from the range's current
// `formulas`. Numbers are computed; formulas are wrapped so they stay live
// (=B2*C2 -> =(B2*C2)*2); text, blanks and booleans are left as they are.
export function arithmetic(formulas: unknown[][], cmd: Extract<Command, { action: "arithmetic" }>): unknown[][] {
  const { op, value } = cmd;
  const f = (n: number) =>
    op === "round" ? roundTo(n, value) : op === "add" ? n + value : op === "subtract" ? n - value : op === "multiply" ? n * value : n / value;
  return formulas.map((row) =>
    row.map((v) => {
      if (typeof v === "number") return f(v);
      if (typeof v === "string" && v.startsWith("="))
        return op === "round" ? `=ROUND(${v.slice(1)},${value})` : `=(${v.slice(1)})${SYM[op]}${value}`;
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
  const places = (d: number) => Number.isInteger(d) && d >= 0 && d <= 10;
  if (cmd.action === "format") {
    const { action, range, ...changes } = cmd;
    if (Object.values(changes).every((v) => v === null)) return `format ${range} names no change`;
    if (cmd.fontSize !== null && !(cmd.fontSize >= 1 && cmd.fontSize <= 409)) return `font size ${cmd.fontSize} is not between 1 and 409`;
    if (cmd.decimals !== null && !places(cmd.decimals)) return `${cmd.decimals} decimal places is not a whole number from 0 to 10`;
  }
  if (cmd.action === "arithmetic" && cmd.op === "round" && !places(cmd.value)) return `rounding to ${cmd.value} places is not a whole number from 0 to 10`;
  if (cmd.action === "set" && typeof cmd.value === "string" && cmd.value.startsWith("="))
    return "typing formulas isn't supported yet";
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
