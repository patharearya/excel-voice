import { describe as group, it, expect } from "vitest";
import { colNum, colName, box, touched, formula, describe, problem, arithmetic, numberFormatCode, textCase, replaceIn, series, matches, answer, readOnly } from "./plan.ts";
import type { Command } from "./commands.ts";

const fmt = (range: string, f: Partial<Extract<Command, { action: "format" }>>): Command => ({
  action: "format", range, bold: null, italic: null, underline: null, strikethrough: null, fontSize: null, fontName: null, align: null, wrap: null, border: null, decimals: null, numberFormat: null, fontColor: null, fillColor: null, ...f,
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

group("batch 2: format options, clear, set, round, more compute", () => {
  it("builds number-format codes with the old defaults and explicit decimals", () => {
    expect(numberFormatCode("currency", null)).toBe("$#,##0.00");
    expect(numberFormatCode("percent", null)).toBe("0%");
    expect(numberFormatCode("number", null)).toBe("#,##0.00");
    expect(numberFormatCode(null, 0)).toBe("#,##0");
    expect(numberFormatCode("percent", 1)).toBe("0.0%");
    expect(numberFormatCode("date", 3)).toBe("m/d/yyyy");
  });
  it("rounds decimally, wraps formulas in ROUND, leaves text", () => {
    const r = { action: "arithmetic", range: "C2:C4", op: "round", value: 2 } as const;
    expect(arithmetic([[1.005], [2.344], ["=A1/3"], ["x"]], r)).toEqual([[1.01], [2.34], ["=ROUND(A1/3,2)"], ["x"]]);
    expect(arithmetic([[2.5]], { ...r, value: 0 })).toEqual([[3]]);
    expect(describe(r)).toBe("Round C2:C4 to 2 decimals");
    expect(problem({ ...r, value: 1.5 })).toMatch(/whole number/);
  });
  it("describes and checks the new format fields", () => {
    expect(describe(fmt("A1:D1", { underline: true, fontSize: 14, align: "center", border: "bottom" }))).toBe(
      "Format A1:D1: underline, size 14, align center, bottom border"
    );
    expect(describe(fmt("C2:C9", { decimals: 1 }))).toBe("Format C2:C9: 1 decimal");
    expect(problem(fmt("A1", { wrap: true }))).toBeNull();
    expect(problem(fmt("A1", { fontSize: 0 }))).toMatch(/between 1 and 409/);
    expect(problem(fmt("A1", { decimals: -1 }))).toMatch(/0 to 10/);
  });
  it("clear and set touch their range and say what they do", () => {
    expect(touched({ action: "clear", range: "E11:E2", what: "contents" })).toBe("E2:E11");
    expect(describe({ action: "clear", range: "A1:D1", what: "formats" })).toBe("Remove formatting from A1:D1");
    expect(describe({ action: "set", range: "B2:B10", value: 0 })).toBe("Put 0 in B2:B10");
    expect(describe({ action: "set", range: "E5", value: "Done" })).toBe('Put "Done" in E5');
    expect(problem({ action: "set", range: "E5", value: "=SUM(A1:A3)" })).toMatch(/formulas/);
  });
  it("maps the new compute ops to Excel functions", () => {
    expect(formula({ action: "compute", op: "counta", range: "A2:A50", target: "A52" })).toBe("=COUNTA(A2:A50)");
    expect(formula({ action: "compute", op: "median", range: "B2:B50", target: "B51" })).toBe("=MEDIAN(B2:B50)");
  });
});

group("batch 3: text, replace, series, highlight, select, ask", () => {
  it("changes text case and spacing, leaving numbers, formulas and blanks", () => {
    const cells = [["  hello   world "], ["MARY o'neil"], [12], ["=A1"], [""]];
    expect(textCase(cells, "upper")).toEqual([["  HELLO   WORLD "], ["MARY O'NEIL"], [12], ["=A1"], [""]]);
    expect(textCase(cells, "proper")).toEqual([["  Hello   World "], ["Mary O'neil"], [12], ["=A1"], [""]]);
    expect(textCase(cells, "trim")).toEqual([["hello world"], ["MARY o'neil"], [12], ["=A1"], [""]]);
  });

  it("replaces whole cells with the typed value and parts of text as text", () => {
    const cmd = { action: "replace", range: "D2:D5", find: "n/a", replacement: 0 } as const;
    expect(replaceIn([["N/A"], ["see N/A note"], ["=IF(A1,\"N/A\",1)"], [5]], cmd)).toEqual([[0], ["see 0 note"], ["=IF(A1,\"N/A\",1)"], [5]]);
    // regex characters in `find` are literal
    expect(replaceIn([["a.b"], ["axb"]], { ...cmd, find: "a.b", replacement: "ok" })).toEqual([["ok"], ["axb"]]);
    expect(problem({ ...cmd, find: "" })).toMatch(/nothing to find/);
  });

  it("fills series down columns, across a single row, and wraps months", () => {
    expect(series({ action: "series", range: "A2:A5", kind: "number", start: 1, step: 1 })).toEqual([[1], [2], [3], [4]]);
    expect(series({ action: "series", range: "B1:E1", kind: "month", start: 11, step: 1 })).toEqual([["Nov", "Dec", "Jan", "Feb"]]);
    expect(series({ action: "series", range: "C1:C3", kind: "number", start: 0.1, step: 0.2 })).toEqual([[0.1], [0.3], [0.5]]);
    expect(series({ action: "series", range: "A1:A2", kind: "weekday", start: 7, step: 1 })).toEqual([["Sun"], ["Mon"]]);
    expect(problem({ action: "series", range: "A1:A9", kind: "number", start: 1, step: 0 })).toMatch(/step of 0/);
  });

  it("highlights only cells that pass the test", () => {
    const h = { action: "highlight", range: "C2:C6", test: "gt", value: 100, fillColor: "#FFFF00", fontColor: null } as const;
    const vals = [[150], [100], ["200"], [""], [99.5]];
    expect(matches(vals, h)).toEqual([[true], [false], [false], [false], [false]]);
    expect(matches(vals, { ...h, test: "lt", value: 0 })).toEqual([[false], [false], [false], [false], [false]]);
    expect(matches([["Apple pie"], ["banana"], [""]], { ...h, test: "contains", value: "APPLE" })).toEqual([[true], [false], [false]]);
    expect(matches([["Done"], ["done"], ["x"]], { ...h, test: "eq", value: "done" })).toEqual([[true], [true], [false]]);
    expect(matches([[1], [""]], { ...h, test: "blank", value: null })).toEqual([[false], [true]]);
    // Empty cells never match "not equal": "highlight what isn't Done" must not paint blanks.
    expect(matches([["Done"], ["Open"], [""]], { ...h, test: "ne", value: "done" })).toEqual([[false], [true], [false]]);
    expect(problem({ ...h, fillColor: null })).toMatch(/no colour/);
    expect(problem({ ...h, value: "lots" })).toMatch(/not a number/);
    expect(describe(h)).toBe("Highlight cells in C2:C6 that are over 100: fill #FFFF00");
  });

  it("answers questions the way Excel's functions would", () => {
    const vals = [[10], [20], ["x"], [""], [30], [40]];
    expect(answer(vals, "sum")).toBe(100);
    expect(answer(vals, "average")).toBe(25);
    expect(answer(vals, "median")).toBe(25);
    expect(answer([[3], [1], [2]], "median")).toBe(2);
    expect(answer(vals, "count")).toBe(4);
    expect(answer(vals, "counta")).toBe(5);
    expect(answer([["a"]], "max")).toBeNull();
    expect(describe({ action: "ask", op: "sum", range: "D2:D50" })).toBe("Total of D2:D50");
  });

  it("marks select and ask as read-only, nothing else", () => {
    expect(readOnly({ action: "select", range: "C1" })).toBe(true);
    expect(readOnly({ action: "ask", op: "max", range: "C2:C9" })).toBe(true);
    expect(readOnly({ action: "set", range: "C1", value: 1 })).toBe(false);
  });

  it("describes strikethrough and font name, and refuses odd font names", () => {
    expect(describe(fmt("A5:D5", { strikethrough: true, fontName: "Arial" }))).toBe("Format A5:D5: strikethrough, font Arial");
    expect(problem(fmt("A1", { fontName: "<script>" }))).toMatch(/not a font name/);
  });
});
