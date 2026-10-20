import { describe, it, expect, vi, afterEach } from "vitest";
import { fetchWithRetry, interpret } from "./interpret.ts";

const ctx = { usedRange: "A1:D50", headers: { A: "Item", B: "Quantity", C: "Price", D: "Total" } };
const ok = (content: unknown) =>
  new Response(JSON.stringify({ choices: [{ message: { content: JSON.stringify(content) } }], usage: {} }), { status: 200 });

afterEach(() => vi.unstubAllGlobals());

describe("fetchWithRetry", () => {
  it("retries a 429 using retry-after and returns the eventual response", async () => {
    const calls: number[] = [];
    vi.stubGlobal("fetch", async () => {
      calls.push(Date.now());
      return calls.length === 1 ? new Response("", { status: 429, headers: { "retry-after": "0" } }) : new Response("ok", { status: 200 });
    });
    const { r } = await fetchWithRetry("x", {});
    expect(r.status).toBe(200);
    expect(calls.length).toBe(2);
  });

  it("retries a thrown network error instead of propagating it", async () => {
    let n = 0;
    vi.stubGlobal("fetch", async () => {
      if (n++ === 0) throw new TypeError("fetch failed", { cause: { code: "UND_ERR_CONNECT_TIMEOUT" } });
      return new Response("ok", { status: 200 });
    });
    const { r } = await fetchWithRetry("x", {}, 1);
    expect(r.status).toBe(200);
  });

  it("gives up after `tries` and lets interpret turn it into a rejection", async () => {
    vi.stubGlobal("fetch", async () => {
      throw new TypeError("fetch failed", { cause: { code: "UND_ERR_CONNECT_TIMEOUT" } });
    });
    await expect(fetchWithRetry("x", {}, 0)).rejects.toThrow("fetch failed");
  });
});

describe("interpret", () => {
  it("returns a validated command list on a good model reply", async () => {
    vi.stubGlobal("fetch", async () => ok({ commands: [{ action: "copy", from: "A2", to: "C3" }] }));
    const r = await interpret("copy A2 to C3", ctx, { apiKey: "k" });
    expect(r.ok).toBe(true);
    if (r.ok) expect(r.commands[0]).toEqual({ action: "copy", from: "A2", to: "C3" });
  });

  it("rejects with a reason when the model reply fails the schema", async () => {
    vi.stubGlobal("fetch", async () => ok({ commands: [{ action: "copy", from: "a2", to: "C3" }] }));
    const r = await interpret("copy a2 to C3", ctx, { apiKey: "k" });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toMatch(/from/);
  });

  it("rejects with a reason on a non-JSON reply", async () => {
    vi.stubGlobal("fetch", async () => new Response(JSON.stringify({ choices: [{ message: { content: "sure!" } }] }), { status: 200 }));
    const r = await interpret("hi", ctx, { apiKey: "k" });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toMatch(/not JSON/);
  });

  it("rejects with a reason on a non-200 from Groq", async () => {
    vi.stubGlobal("fetch", async () => new Response(JSON.stringify({ error: "bad" }), { status: 500 }));
    const r = await interpret("hi", ctx, { apiKey: "k" });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toMatch(/groq 500/);
  });
});
