// Transcript + sheet context -> validated command list, via a Groq LLM.
// The model only interprets; nothing here touches a sheet.
import { parseCommands, responseJsonSchema, type ParseResult } from "./commands.ts";

export type SheetContext = {
  usedRange: string; // e.g. "A1:D50"
  headers: Record<string, string>; // column letter -> header text, from row 1
};

export const DEFAULT_MODEL = "openai/gpt-oss-120b";

// The mechanism, not just the preference: the model needs to know how ranges
// are derived from the context, or its own reasonable reading will differ.
export function systemPrompt(ctx: SheetContext): string {
  const headers = Object.entries(ctx.headers)
    .map(([col, name]) => `${col}="${name}"`)
    .join(", ");
  const [first, last = first] = ctx.usedRange.split(":"); // an empty sheet's used range is a single cell
  const lastRow = last.replace(/[A-Z]+/, "");
  const firstRow = first.replace(/[A-Z]+/, "");
  return `You turn one spoken instruction about an Excel sheet into a list of commands.
Output JSON only, matching the schema. Each utterance may need several commands, in the order spoken.

Sheet context:
- used range: ${ctx.usedRange} (rows ${firstRow} to ${lastRow})
- row 1 is the header row. headers: ${headers}

How to resolve ranges:
- A named column ("the price column") means the column whose header matches: look it up in the headers.
- "column C" or a named column with no rows given means the data rows only: C2:C${lastRow}. Exception: copying a whole column includes the header: C1:C${lastRow}.
- "the header row" means row 1 across the used range: ${first}:${last.replace(/[0-9]+/, "")}1.
- "the table" or "everything" means the used range ${ctx.usedRange}.
- "at the bottom of column D" means the first empty row: D${Number(lastRow) + 1}.
- Spoken ranges: "A2 through A10", "A2 to A10", "A2 and A10" (a mishearing of "A2 to A10" when both are cells) all mean A2:A10 for format/sort/compute. For copy, "copy A2 to C3" means from A2 to C3.
- Cells and ranges use uppercase letters, no sheet names, no dollar signs.

Commands:
- copy: from (range), to (top-left destination cell).
- format: range, then any of bold / italic / underline / strikethrough (true|false), fontSize (points), fontName (e.g. Arial), align (left|center|right), wrap (true|false), border (all = every cell boxed, outline = around the outside, bottom = a cell border along the bottom edge, none = remove). Two different things: underline marks the text itself ("underline the header row"); a line, rule or border under/below cells is border bottom ("add a line under the header row", "put a line below row 1"), numberFormat (currency|percent|number|date), decimals (places after the point, e.g. "2 decimals" = 2; alone it means numberFormat number), fontColor / fillColor. Set only what was asked; every other field null. Colors as 6-digit hex, e.g. red #FF0000, yellow #FFFF00, green #00FF00, blue #0000FF.
- compute: op (sum|average|min|max|count|counta|median), range, target cell. "total" means sum. count counts numbers; counta counts non-empty cells ("how many filled", "count the entries"). It writes a formula into the target.
- arithmetic: change the numbers already in a range, in place: range, op (add|subtract|multiply|divide|round), value. "double" = multiply 2, "triple" = multiply 3, "halve" = divide 2, "increase by 10%" = multiply 1.1, "reduce by 10%" = multiply 0.9, "round to 2 decimals" = round 2, "round to whole numbers" = round 0. It does not write results anywhere else and cannot recalculate other columns from the change.
- clear: range, what (contents = delete values/formulas but keep formatting; formats = remove formatting, keep values; all = both). "clear"/"empty"/"delete the values in" = contents.
- set: range, value (a number, or text exactly as spoken). Puts the same value in every cell of the range: "put 0 in B2 to B10", "write Done in E5". Never a formula.
- text: range, op (upper = "uppercase"/"capitals", lower = "lowercase", proper = "capitalize"/"title case"/"capitalize each word", trim = "remove extra spaces"/"tidy spaces").
- replace: range, find (text as spoken), replacement (a number if it is a number, else text). "replace N/A with 0 in column D". With no range given, use the used range.
- series: range, kind (number|month|weekday), start, step. Fills each column top to bottom (or a single row left to right). number: start = first value, step = gap (1 unless said). month: start 1 = January ... 12 = December. weekday: start 1 = Monday ... 7 = Sunday. "number A2 to A20 from 1" = number start 1 step 1; "fill B2 to B13 with Jan to Dec" = month start 1 step 1.
- highlight: colour only the cells in a range that pass a test: range, test (gt = over/more than/above, ge = at least, lt = under/less than/below, le = at most, eq = equal to/is, ne = not equal, contains = contains/includes, blank = empty), value (a number for gt/ge/lt/le; the text or number for eq/ne/contains; null for blank), fillColor and/or fontColor. "highlight ... in yellow" = fillColor; "make ... red"/"color the text red" = fontColor. "negative" = lt 0. Use highlight whenever only some cells should change colour; use format when all of them should.
- select: range. "select column C", "go to D50", "show me the header row". Changes nothing.
- ask: op (sum|average|min|max|count|counta|median), range. A question about the sheet ("what's the total of column D?", "how many items are there?", "what's the highest price?") is ask: the answer is shown in the panel and nothing is written. Use compute only when told where to put the result.
- sort: range (the used range unless told otherwise), by (column letter), order (asc|desc; "high to low"/"largest first" = desc, "alphabetically"/"A to Z" = asc), hasHeader true when the range starts at row 1.

If the instruction cannot be expressed with these commands, or is not an instruction at all (greetings, "thank you", silence), return {"commands": []}. Never invent a command to fill the gap.`;
}

export type InterpretResult = ParseResult & { model: string; ms: number; usage?: unknown };

// Groq free tier rate-limits per minute. Honour retry-after on 429 and retry
// once on a network error, up to `tries` times, then give up. `ms` is the
// final attempt only, so latency reports are not inflated by waiting.
export async function fetchWithRetry(url: string, init: RequestInit, tries = 3): Promise<{ r: Response; ms: number }> {
  for (let i = 0; ; i++) {
    const t0 = Date.now();
    let r: Response;
    try {
      r = await fetch(url, init);
    } catch (e: any) {
      if (i >= tries) throw e;
      console.warn(`  network error (${e?.cause?.code ?? e?.message}), retry ${i + 1}/${tries} in 3000 ms`);
      await new Promise((res) => setTimeout(res, 3000));
      continue;
    }
    if (r.status !== 429 || i >= tries) return { r, ms: Date.now() - t0 };
    const wait = Math.min(Number(r.headers.get("retry-after") ?? 2) * 1000 || 2000, 30000);
    console.warn(`  groq 429, retry ${i + 1}/${tries} in ${wait} ms`);
    await new Promise((res) => setTimeout(res, wait));
  }
}

export async function interpret(
  transcript: string,
  ctx: SheetContext,
  opts: { model?: string; apiKey?: string } = {}
): Promise<InterpretResult> {
  const model = opts.model ?? DEFAULT_MODEL;
  const apiKey = opts.apiKey ?? process.env.GROQ_API_KEY;
  if (!apiKey) return { ok: false, reason: "GROQ_API_KEY not set", raw: null, model, ms: 0 };
  // Built outside the try below, so a bad context throws to the caller
  // instead of being mislabelled as a network error.
  const system = systemPrompt(ctx);
  let r: Response;
  let ms = 0;
  try {
    ({ r, ms } = await fetchWithRetry("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      model,
      temperature: 0,
      messages: [
        { role: "system", content: system },
        { role: "user", content: transcript },
      ],
      response_format: {
        type: "json_schema",
        json_schema: { name: "commands", strict: true, schema: responseJsonSchema },
      },
    }),
    }));
  } catch (e: any) {
    return { ok: false, reason: `network: ${e?.cause?.code ?? e?.message ?? e}`, raw: null, model, ms };
  }
  const body: any = await r.json().catch(() => ({}));
  if (!r.ok) return { ok: false, reason: `groq ${r.status}: ${JSON.stringify(body).slice(0, 300)}`, raw: body, model, ms };
  const text = body.choices?.[0]?.message?.content ?? "";
  let json: unknown;
  try {
    json = JSON.parse(text);
  } catch {
    return { ok: false, reason: `not JSON: ${text.slice(0, 200)}`, raw: text, model, ms, usage: body.usage };
  }
  return { ...parseCommands(json), model, ms, usage: body.usage };
}
