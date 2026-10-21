import { describe as group, it, expect } from "vitest";
import { colNum, colName, box, touched, formula, describe, problem, arithmetic } from "./plan.ts";
import type { Command } from "./commands.ts";

const fmt = (range: string, f: Partial<Extract<Command, { action: "format" }>>): Command => ({
  action: "format", range, bold: null, italic: null, numberFormat: null, fontColor: null, fillColor: null, ...f,
});

group("columns", () => {
  it("round-trips letters and numbers across the Z/AA and ZZ/AAA boundaries", () => {
    for (const [s, n] of [["A", 1], ["Z", 26], ["AA", 27], ["AZ", 52], ["ZZ", 702], ["AAA", 703], ["XFD", 16384]] as const) {
      expect(colNum(s)).toBe(n);
      expect(colName(n)).toBe(s);
    }
  });
  it("normalises a reversed range", () => {
    expect(box("C50:A2")).toEqual({ c1: 1, r1: 2, c2: 3, r2: 50 });
  });
});

group("touched", () => {
  it("copy covers the destination at the source's size", () => {
    expect(touched({ action: "copy", from: "C1:C50", to: "H1" })).toBe("H1:H50");
    expect(touched({ action: "copy", from: "A2:B3", to: "Y10" })).toBe("Y10:Z11");
    expect(touched({ action: "copy", from: "A2", to: "C3" })).toBe("C3");
  });
  it("compute touches only the target; sort and format the range", () => {
    expect(touched({ action: "compute", op: "sum", range: "D2:D50", target: "D51" })).toBe("D51");
    expect(touched({ action: "sort", range: "A1:D50", by: "C", order: "desc", hasHeader: true })).toBe("A1:D50");
    expect(touched(fmt("D1:A1", { bold: true }))).toBe("A1:D1");
  });
});

group("preview text", () => {
  it("states the formula and the format changes", () => {
    expect(formula({ action: "compute", op: "average", range: "B2:B50", target: "B51" })).toBe("=AVERAGE(B2:B50)");
    expect(describe(fmt("C2:C50", { numberFormat: "currency", bold: true }))).toBe("Format C2:C50: bold, currency");
    expect(describe({ action: "copy", from: "C1:C50", to: "H1" })).toBe("Copy C1:C50 to H1:H50");
  });
});

group("arithmetic", () => {
  const dbl = { action: "arithmetic", range: "B2:B4", op: "multiply", value: 2 } as const;
  it("computes numbers, wraps formulas, leaves text and blanks", () => {
    expect(arithmetic([[7], ["=B2*C2"], ["n/a"], [""], [true]], dbl)).toEqual([[14], ["=(B2*C2)*2"], ["n/a"], [""], [true]]);
    expect(arithmetic([[10]], { ...dbl, op: "subtract", value: 3 })).toEqual([[7]]);
    expect(arithmetic([["=A1"]], { ...dbl, op: "divide", value: 4 })).toEqual([["=(A1)/4"]]);
  });
  it("previews in plain words and refuses division by zero", () => {
    expect(describe(dbl)).toBe("Multiply B2:B4 by 2");
    expect(describe({ ...dbl, op: "add", value: 10 })).toBe("Add 10 to B2:B4");
    expect(touched(dbl)).toBe("B2:B4");
    expect(problem({ ...dbl, op: "divide", value: 0 })).toMatch(/zero/);
    expect(problem(dbl)).toBeNull();
  });
});

group("problem", () => {
  it("passes ordinary commands", () => {
    expect(problem({ action: "compute", op: "sum", range: "D2:D50", target: "D51" })).toBeNull();
    expect(problem({ action: "sort", range: "A1:D50", by: "C", order: "asc", hasHeader: true })).toBeNull();
    expect(problem(fmt("A1:D1", { bold: true }))).toBeNull();
  });
  it("rejects a compute target inside its own range", () => {
    expect(problem({ action: "compute", op: "sum", range: "D2:D50", target: "D50" })).toMatch(/circular/);
  });
  it("rejects a sort column outside the range", () => {
    expect(problem({ action: "sort", range: "A1:D50", by: "F", order: "asc", hasHeader: true })).toMatch(/not in/);
  });
  it("rejects a format with no change", () => {
    expect(problem(fmt("A1", {}))).toMatch(/no change/);
  });
  it("rejects a copy that runs off the sheet", () => {
    expect(problem({ action: "copy", from: "A1:A10", to: "A1048570" })).toMatch(/outside/);
    expect(problem({ action: "copy", from: "A1:C1", to: "XFC1" })).toMatch(/outside/);
  });
});
