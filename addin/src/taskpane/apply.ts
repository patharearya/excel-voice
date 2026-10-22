/* global Excel */
// Runs validated commands on the active sheet. Excel's own Ctrl+Z does not undo
// add-in changes, so every command snapshots the cells it will change first and
// Undo writes the snapshots back in reverse order.
import type { Command } from "../../../proxy/commands.ts";
import {
  touched, describe, formula, size, box, address, colNum, arithmetic, numberFormatCode,
  textCase, replaceIn, series, matches, answer,
} from "../../../proxy/plan.ts";

const EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"] as const;

// What a snapshot keeps. ponytail: conditional formats, comments, merges, data
// validation and spilled arrays are not restored; add them if Undo is seen to miss one.
const LOAD: Excel.CellPropertiesLoadOptions = {
  format: {
    font: { bold: true, italic: true, color: true, name: true, size: true, underline: true, strikethrough: true },
    fill: { color: true, pattern: true },
    borders: { style: true, color: true, weight: true },
    horizontalAlignment: true,
    verticalAlignment: true,
    wrapText: true,
  },
};

// `values` is read for commands that test what cells show (highlight); Undo restores `formulas`.
type Snap = { address: string; formulas: any[][]; values: any[][]; numberFormat: any[][]; props: Excel.CellProperties[][] };
export type Undo = { sheetId: string; snaps: Snap[]; done: string[] };

export type SheetContext = { usedRange: string; headers: Record<string, string> };

// Used range and row-1 headers, for /interpret to resolve "the price column".
export async function sheetContext(): Promise<SheetContext> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("address");
    await ctx.sync();
    if (used.isNullObject) return { usedRange: "A1", headers: {} };
    const b = box(used.address.split("!").pop()!.replace(/\$/g, ""));
    const row1 = sheet.getRange(address({ ...b, r1: 1, r2: 1 }));
    row1.load("values");
    await ctx.sync();
    const headers: Record<string, string> = {};
    row1.values[0].forEach((v, i) => {
      if (v !== "") headers[address({ c1: b.c1 + i, r1: 1, c2: b.c1 + i, r2: 1 }).replace(/[0-9]+/, "")] = String(v);
    });
    return { usedRange: address(b), headers };
  });
}

// Preview: select every cell the commands would change. Nothing changes yet.
export async function select(cmds: Command[]) {
  await Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getActiveWorksheet().getRanges(cmds.map(touched).join(",")).select();
    await ctx.sync();
  });
}

// Answers for ask commands, computed from the cells' current values. Nothing is written.
export async function ask(cmds: Extract<Command, { action: "ask" }>[]): Promise<string[]> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const ranges = cmds.map((c) => sheet.getRange(touched(c)).load("values"));
    await ctx.sync();
    return cmds.map((c, i) => {
      const a = answer(ranges[i].values, c.op);
      return `${describe(c)}: ${a === null ? "no numbers there" : a.toLocaleString(undefined, { maximumFractionDigits: 4 })}`;
    });
  });
}

// Runs commands in order. On a failure it stops and returns the error together
// with the snapshots taken so far, so the commands that did run can still be undone.
export async function apply(cmds: Command[]): Promise<{ undo: Undo; error?: string }> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load("id");
    await ctx.sync();
    const undo: Undo = { sheetId: sheet.id, snaps: [], done: [] };
    for (const cmd of cmds) {
      try {
        if (cmd.action === "select" || cmd.action === "ask") {
          // Changes nothing: no snapshot, nothing for Undo.
          sheet.getRange(touched(cmd)).select();
          await ctx.sync();
          continue;
        }
        const snap = await snapshot(ctx, sheet, touched(cmd));
        undo.snaps.push(snap);
        run(sheet, cmd, snap);
        await ctx.sync();
        undo.done.push(describe(cmd));
      } catch (e: any) {
        return { undo, error: `${describe(cmd)}: ${e?.message ?? e}` };
      }
    }
    return { undo };
  });
}

export async function undo(u: Undo) {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(u.sheetId);
    for (const s of [...u.snaps].reverse()) restore(sheet, s);
    await ctx.sync();
  });
}

async function snapshot(ctx: Excel.RequestContext, sheet: Excel.Worksheet, addr: string): Promise<Snap> {
  const r = sheet.getRange(addr);
  r.load("formulas,values,numberFormat");
  const props = r.getCellProperties(LOAD);
  await ctx.sync();
  return { address: addr, formulas: r.formulas, values: r.values, numberFormat: r.numberFormat, props: props.value };
}

function restore(sheet: Excel.Worksheet, s: Snap) {
  const r = sheet.getRange(s.address);
  r.formulas = s.formulas; // constants come back too: formulas holds the value when there is no formula
  r.numberFormat = s.numberFormat;
  // A cell with no fill reads back as white; writing white would add a fill.
  // Clear all fills, then write back only the real ones.
  r.format.fill.clear();
  r.setCellProperties(
    s.props.map((row) =>
      row.map(({ format }) => {
        const { fill, ...rest } = format ?? {};
        return { format: fill?.pattern === "None" ? rest : { ...rest, fill } };
      })
    ) as Excel.SettableCellProperties[][]
  );
}

// Inside lines only exist when the range has more than one row or column;
// setting them on a single cell or line is skipped rather than risked.
function border(r: Excel.Range, kind: "all" | "outline" | "bottom" | "none", { rows, cols }: { rows: number; cols: number }) {
  const set = (edge: string, on: boolean) => {
    r.format.borders.getItem(edge as Excel.BorderIndex).style = on ? "Continuous" : "None";
  };
  const inside = [...(rows > 1 ? ["InsideHorizontal"] : []), ...(cols > 1 ? ["InsideVertical"] : [])];
  if (kind === "bottom") return set("EdgeBottom", true);
  for (const e of EDGES) set(e, kind !== "none");
  for (const e of inside) set(e, kind === "all");
}

// `snap` is the touched range as it was just before this command.
function run(sheet: Excel.Worksheet, cmd: Command, snap: Snap) {
  const r = sheet.getRange(touched(cmd));
  switch (cmd.action) {
    case "copy":
      r.copyFrom(sheet.getRange(cmd.from), Excel.RangeCopyType.all);
      break;
    case "format":
      if (cmd.bold !== null) r.format.font.bold = cmd.bold;
      if (cmd.italic !== null) r.format.font.italic = cmd.italic;
      if (cmd.underline !== null) r.format.font.underline = cmd.underline ? "Single" : "None";
      if (cmd.strikethrough !== null) r.format.font.strikethrough = cmd.strikethrough;
      if (cmd.fontName) r.format.font.name = cmd.fontName;
      if (cmd.fontSize !== null) r.format.font.size = cmd.fontSize;
      if (cmd.align) r.format.horizontalAlignment = ({ left: "Left", center: "Center", right: "Right" } as const)[cmd.align];
      if (cmd.wrap !== null) r.format.wrapText = cmd.wrap;
      if (cmd.border) border(r, cmd.border, size(cmd.range));
      if (cmd.fontColor) r.format.font.color = cmd.fontColor;
      if (cmd.fillColor) r.format.fill.color = cmd.fillColor;
      if (cmd.numberFormat || cmd.decimals !== null) {
        const { rows, cols } = size(cmd.range);
        const code = numberFormatCode(cmd.numberFormat, cmd.decimals);
        r.numberFormat = Array.from({ length: rows }, () => Array(cols).fill(code));
      }
      break;
    case "clear":
      r.clear(({ contents: "Contents", formats: "Formats", all: "All" } as const)[cmd.what]);
      break;
    case "text":
      r.formulas = textCase(snap.formulas, cmd.op) as any[][];
      break;
    case "replace":
      r.formulas = replaceIn(snap.formulas, cmd) as any[][];
      break;
    case "series":
      r.values = series(cmd);
      break;
    case "highlight":
      matches(snap.values, cmd).forEach((row, i) =>
        row.forEach((hit, j) => {
          if (!hit) return;
          const cell = r.getCell(i, j);
          if (cmd.fillColor) cell.format.fill.color = cmd.fillColor;
          if (cmd.fontColor) cell.format.font.color = cmd.fontColor;
        })
      );
      break;
    case "select":
    case "ask":
      break; // handled in apply(), never snapshotted
    case "set": {
      const { rows, cols } = size(cmd.range);
      r.values = Array.from({ length: rows }, () => Array(cols).fill(cmd.value));
      break;
    }
    case "compute":
      r.formulas = [[formula(cmd)]];
      break;
    case "arithmetic":
      r.formulas = arithmetic(snap.formulas, cmd) as any[][];
      break;
    case "sort":
      r.sort.apply([{ key: colNum(cmd.by) - box(cmd.range).c1, ascending: cmd.order === "asc" }], false, cmd.hasHeader);
      break;
  }
}
