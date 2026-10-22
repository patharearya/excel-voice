// Single source of truth for what the LLM may return and what the add-in may run.
// Every field is required-or-null (no optionals) so the same schema works for
// Groq strict structured output, which demands every property be listed.
import { z } from "zod";

const CELL = /^[A-Z]{1,3}[0-9]{1,7}$/;
const RANGE = /^[A-Z]{1,3}[0-9]{1,7}(:[A-Z]{1,3}[0-9]{1,7})?$/;
const COLUMN = /^[A-Z]{1,3}$/;
const HEX = /^#[0-9A-Fa-f]{6}$/;

export const Cell = z.string().regex(CELL, "cell like D51");
export const Range = z.string().regex(RANGE, "range like A1 or C2:C50");
export const Column = z.string().regex(COLUMN, "column letter like C");

export const Copy = z.object({
  action: z.literal("copy"),
  from: Range,
  to: Cell,
});

export const Format = z.object({
  action: z.literal("format"),
  range: Range,
  bold: z.boolean().nullable(),
  italic: z.boolean().nullable(),
  underline: z.boolean().nullable(),
  strikethrough: z.boolean().nullable(),
  fontSize: z.number().nullable(),
  fontName: z.string().nullable(),
  align: z.enum(["left", "center", "right"]).nullable(),
  wrap: z.boolean().nullable(),
  border: z.enum(["all", "outline", "bottom", "none"]).nullable(),
  numberFormat: z.enum(["currency", "percent", "number", "date"]).nullable(),
  decimals: z.number().nullable(), // places after the point; with no numberFormat it means "number"
  fontColor: z.string().regex(HEX, "hex like #FF0000").nullable(),
  fillColor: z.string().regex(HEX, "hex like #FFFF00").nullable(),
});

export const Compute = z.object({
  action: z.literal("compute"),
  op: z.enum(["sum", "average", "min", "max", "count", "counta", "median"]),
  range: Range,
  target: Cell,
});

export const Sort = z.object({
  action: z.literal("sort"),
  range: Range,
  by: Column,
  order: z.enum(["asc", "desc"]),
  hasHeader: z.boolean(),
});

// Change numbers in place: "double B5", "add 10 to column C", "round C to 2
// decimals" (op round, value = decimal places).
export const Arithmetic = z.object({
  action: z.literal("arithmetic"),
  range: Range,
  op: z.enum(["add", "subtract", "multiply", "divide", "round"]),
  value: z.number(),
});

// Empty cells: their contents, their formatting, or both.
export const Clear = z.object({
  action: z.literal("clear"),
  range: Range,
  what: z.enum(["contents", "formats", "all"]),
});

// Type one value into every cell of a range: "put 0 in B2 to B10".
export const SetValue = z.object({
  action: z.literal("set"),
  range: Range,
  value: z.union([z.string(), z.number()]),
});

// Change text in place: case, or tidy spaces.
export const TextCase = z.object({
  action: z.literal("text"),
  range: Range,
  op: z.enum(["upper", "lower", "proper", "trim"]),
});

// Find and replace inside a range. A cell equal to `find` becomes `replacement`;
// inside longer text, the matching part is replaced.
export const Replace = z.object({
  action: z.literal("replace"),
  range: Range,
  find: z.string(),
  replacement: z.union([z.string(), z.number()]),
});

// Fill a range with a series: 1, 2, 3... or Jan, Feb... or Mon, Tue...
export const Series = z.object({
  action: z.literal("series"),
  range: Range,
  kind: z.enum(["number", "month", "weekday"]),
  start: z.number(), // number: first value; month: 1 = Jan; weekday: 1 = Mon
  step: z.number(),
});

// Colour only the cells that pass a test. The test is evaluated in code.
export const Highlight = z.object({
  action: z.literal("highlight"),
  range: Range,
  test: z.enum(["gt", "ge", "lt", "le", "eq", "ne", "contains", "blank"]),
  value: z.union([z.string(), z.number()]).nullable(),
  fillColor: z.string().regex(HEX, "hex like #FFFF00").nullable(),
  fontColor: z.string().regex(HEX, "hex like #FF0000").nullable(),
});

// Read-only: select cells, or answer a question about them in the panel.
export const Select = z.object({
  action: z.literal("select"),
  range: Range,
});
export const Ask = z.object({
  action: z.literal("ask"),
  op: z.enum(["sum", "average", "min", "max", "count", "counta", "median"]),
  range: Range,
});

export const Command = z.discriminatedUnion("action", [
  Copy, Format, Compute, Sort, Arithmetic, Clear, SetValue, TextCase, Replace, Series, Highlight, Select, Ask,
]);
export const Response = z.object({ commands: z.array(Command) });

export type Command = z.infer<typeof Command>;
export type Response = z.infer<typeof Response>;

// JSON schema handed to Groq. Zod 4 emits additionalProperties:false and a full
// required list for z.object, which is what strict mode needs.
export const responseJsonSchema = z.toJSONSchema(Response);

export type ParseResult = { ok: true; commands: Command[] } | { ok: false; reason: string; raw: unknown };

// Validate anything (LLM output, test input). Never throws; a failure carries a reason.
export function parseCommands(raw: unknown): ParseResult {
  const r = Response.safeParse(raw);
  if (r.success) return { ok: true, commands: r.data.commands };
  const reason = r.error.issues.map((i) => `${i.path.join(".") || "(root)"}: ${i.message}`).join("; ");
  return { ok: false, reason, raw };
}
