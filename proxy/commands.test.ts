import { describe, it, expect } from "vitest";
import { parseCommands, responseJsonSchema } from "./commands.ts";

const ok = (commands: unknown[]) => parseCommands({ commands });

describe("command schema", () => {
  it("accepts every example shape from the brief", () => {
    const r = ok([
      { action: "copy", from: "C2:C50", to: "H2" },
      { action: "format", range: "A1:D1", bold: true, italic: null, numberFormat: null, fontColor: null, fillColor: null },
      { action: "format", range: "C2:C5", bold: null, italic: null, numberFormat: "currency", fontColor: null, fillColor: null },
      { action: "compute", op: "sum", range: "D2:D49", target: "D50" },
      { action: "sort", range: "A1:D50", by: "C", order: "desc", hasHeader: true },
    ]);
    expect(r.ok).toBe(true);
    if (r.ok) expect(r.commands).toHaveLength(5);
  });

  it("accepts an empty list (the 'did not understand' answer)", () => {
    expect(ok([]).ok).toBe(true);
  });

  it.each([
    ["unknown action", { action: "delete", range: "B1:B50" }],
    ["lowercase range", { action: "copy", from: "c2:c50", to: "H2" }],
    ["sheet-qualified range", { action: "copy", from: "Sheet1!C2:C50", to: "H2" }],
    ["copy target is a range, not a cell", { action: "copy", from: "C2:C50", to: "H2:H50" }],
    ["bad number format", { action: "format", range: "C2:C5", bold: null, italic: null, numberFormat: "money", fontColor: null, fillColor: null }],
    ["colour name instead of hex", { action: "format", range: "C2:C5", bold: null, italic: null, numberFormat: null, fontColor: "red", fillColor: null }],
    ["bad op", { action: "compute", op: "median", range: "D2:D49", target: "D50" }],
    ["sort by a cell instead of a column", { action: "sort", range: "A1:D50", by: "C1", order: "desc", hasHeader: true }],
    ["missing hasHeader", { action: "sort", range: "A1:D50", by: "C", order: "desc" }],
  ])("rejects %s with a reason", (_name, cmd) => {
    const r = ok([cmd]);
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason.length).toBeGreaterThan(0);
  });

  it("rejects non-object input without throwing", () => {
    expect(parseCommands("yes").ok).toBe(false);
    expect(parseCommands(null).ok).toBe(false);
    expect(parseCommands({ commands: "none" }).ok).toBe(false);
  });

  it("emits a strict-mode JSON schema for Groq (all keys required, no extras)", () => {
    const s: any = responseJsonSchema;
    expect(s.required).toEqual(["commands"]);
    expect(s.additionalProperties).toBe(false);
    const variants: any[] = s.properties.commands.items.anyOf ?? s.properties.commands.items.oneOf;
    expect(variants.length).toBe(4);
    for (const v of variants) {
      expect(v.additionalProperties).toBe(false);
      expect(v.required.sort()).toEqual(Object.keys(v.properties).sort());
    }
  });
});
