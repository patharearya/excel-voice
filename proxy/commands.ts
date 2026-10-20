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
  numberFormat: z.enum(["currency", "percent", "number"]).nullable(),
  fontColor: z.string().regex(HEX, "hex like #FF0000").nullable(),
  fillColor: z.string().regex(HEX, "hex like #FFFF00").nullable(),
});

export const Compute = z.object({
  action: z.literal("compute"),
  op: z.enum(["sum", "average", "min", "max", "count"]),
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

export const Command = z.discriminatedUnion("action", [Copy, Format, Compute, Sort]);
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
