import { describe, expect, it, vi } from "vitest";
import { existsSync, readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { createRequire } from "node:module";
const { JSDOM } = createRequire(import.meta.url)("jsdom");
import path from "node:path";

const sourcePath = path.resolve("src/taskpane/theme-bootstrap.js");
const html = readFileSync("src/taskpane/taskpane.html", "utf8");
function host(stored?: string, officeTheme?: unknown, dark = false) {
  const dom = new JSDOM(html, { url: "https://localhost:3001", runScripts: "outside-only" });
  const win = dom.window as any;
  const change = vi.fn();
  win.matchMedia = () => ({ matches: dark, addEventListener: change });
  win.Office = { context: { officeTheme } };
  if (stored !== undefined) win.localStorage.setItem("omnia.theme", stored);
  if (existsSync(sourcePath)) win.eval(readFileSync(sourcePath, "utf8"));
  return { win, api: win.OmniaTheme, doc: win.document, change };
}

describe("Omnia theme resolution and persistence", () => {
  it("defaults to Teal + Auto before paint", () => {
    const { api, doc } = host();
    expect(api.load()).toEqual({ palette: "teal", mode: "auto" });
    expect(doc.documentElement.dataset.palette).toBe("teal");
    expect(doc.documentElement.dataset.scheme).toBe("light");
  });
  it.each(["teal", "blue", "gold", "red"])("restores stored %s palette and explicit mode", palette => {
    const { api, doc } = host(JSON.stringify({ palette, mode: "dark" }), { isDarkTheme: false });
    expect(api.load()).toEqual({ palette, mode: "dark" });
    expect(doc.documentElement.dataset.palette).toBe(palette);
    expect(doc.documentElement.dataset.scheme).toBe("dark");
  });
  it.each(["not json", "null", "[]", '{"palette":"purple","mode":"bad"}'])("handles invalid storage %s", stored => {
    expect(host(stored).api.load()).toEqual({ palette: "teal", mode: "auto" });
  });
  it("normalizes palette and mode independently", () => {
    expect(host('{"palette":"red","mode":"bad"}').api.load()).toEqual({ palette: "red", mode: "auto" });
  });
  it.each([true, false])("Auto follows Office isDarkTheme=%s before color or browser", dark => {
    const { doc } = host(undefined, { isDarkTheme: dark, bodyBackgroundColor: dark ? "#ffffff" : "#000000" }, !dark);
    expect(doc.documentElement.dataset.scheme).toBe(dark ? "dark" : "light");
  });
  it.each([["#202020", "dark"], ["#ffffff", "light"], ["000000", "dark"], ["#fff", "light"], ["#808080", "dark"]])("Auto resolves Office luminance %s", (color, scheme) => {
    expect(host(undefined, { bodyBackgroundColor: color }).doc.documentElement.dataset.scheme).toBe(scheme);
  });
  it.each([true, false])("falls back to browser dark=%s with unavailable Office theme", dark => {
    expect(host(undefined, undefined, dark).doc.documentElement.dataset.scheme).toBe(dark ? "dark" : "light");
    expect(host(undefined, { bodyBackgroundColor: "invalid" }, dark).doc.documentElement.dataset.scheme).toBe(dark ? "dark" : "light");
  });
  it("refreshes Auto when Office becomes ready or its theme changes", () => {
    const { api, win, doc } = host();
    win.Office.context.officeTheme = { isDarkTheme: true };
    api.refresh();
    expect(doc.documentElement.dataset.scheme).toBe("dark");
    win.Office.context.officeTheme = { isDarkTheme: false };
    api.refresh();
    expect(doc.documentElement.dataset.scheme).toBe("light");
  });
  it("persists exactly palette and mode and applies data attributes", () => {
    const { api, win, doc } = host();
    api.save({ palette: "gold", mode: "light" });
    expect(JSON.parse(win.localStorage.getItem("omnia.theme"))).toEqual({ palette: "gold", mode: "light" });
    expect(doc.documentElement.dataset.palette).toBe("gold");
    expect(doc.documentElement.dataset.scheme).toBe("light");
  });
  it("survives denied reads/writes and unavailable media queries", () => {
    const { api, win, doc } = host();
    Object.defineProperty(win, "localStorage", { get() { throw new Error("denied"); } });
    win.matchMedia = undefined;
    expect(api.load()).toEqual({ palette: "teal", mode: "auto" });
    expect(() => api.save({ palette: "red", mode: "auto" })).not.toThrow();
    expect(doc.documentElement.dataset.palette).toBe("red");
  });
  it("survives Office theme access throwing", () => {
    const { api, win, doc } = host(undefined, undefined, true);
    Object.defineProperty(win.Office.context, "officeTheme", { get() { throw new Error("not ready"); } });
    expect(() => api.refresh()).not.toThrow();
    expect(doc.documentElement.dataset.scheme).toBe("dark");
  });
});

describe("Omnia settings controls", () => {
  it("binds palette and appearance controls without submitting connection settings", () => {
    const { api, doc, win } = host();
    api.bindControls();
    const submit = vi.fn();
    doc.querySelector("form")!.addEventListener("submit", submit);
    doc.querySelector('[data-set-palette="blue"]')!.click();
    doc.querySelector('[data-set-mode="dark"]')!.click();
    expect(submit).not.toHaveBeenCalled();
    expect(JSON.parse(win.localStorage.getItem("omnia.theme"))).toEqual({ palette: "blue", mode: "dark" });
    expect(doc.querySelector('[data-set-palette="blue"]')!.getAttribute("aria-checked")).toBe("true");
    expect(doc.querySelector('[data-set-palette="teal"]')!.getAttribute("aria-checked")).toBe("false");
    expect(doc.querySelector('[data-set-mode="dark"]')!.getAttribute("aria-checked")).toBe("true");
    expect(doc.documentElement.dataset.scheme).toBe("dark");
  });
  it("supports keyboard radio navigation and a single tab stop per group", () => {
    const { api, doc, win } = host();
    api.bindControls();
    const teal = doc.querySelector('[data-set-palette="teal"]')!;
    teal.focus();
    teal.dispatchEvent(new win.KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    expect(doc.activeElement!.dataset.setPalette).toBe("blue");
    expect(doc.querySelectorAll('[data-set-palette][tabindex="0"]')).toHaveLength(1);
    doc.activeElement!.dispatchEvent(new win.KeyboardEvent("keydown", { key: "End", bubbles: true }));
    expect(doc.activeElement!.dataset.setPalette).toBe("red");
  });
  it("listens for browser changes but keeps explicit appearance authoritative", () => {
    const { api, change, doc, win } = host();
    api.bindControls();
    expect(change).toHaveBeenCalledWith("change", expect.any(Function));
    api.save({ palette: "teal", mode: "light" });
    win.Office.context.officeTheme = { isDarkTheme: true };
    change.mock.calls[0][1]();
    expect(doc.documentElement.dataset.scheme).toBe("light");
  });
});

describe("shared kit and pre-paint contracts", () => {
  it("vendors the exact approved rev-3 token bytes", () => {
    const tokens = readFileSync("src/taskpane/omnia-tokens.css");
    expect(createHash("sha256").update(tokens).digest("hex")).toBe("9d4a6a258b587d6c65c9773521ef84e3442915027e052d1f2e82725242cfc3d5");
  });
  it("loads the synchronous theme bootstrap and tokens before taskpane styles", () => {
    expect(html.indexOf('src="theme-bootstrap.js"')).toBeGreaterThan(0);
    expect(html.indexOf('src="theme-bootstrap.js"')).toBeLessThan(html.indexOf('href="omnia-tokens.css"'));
    expect(html.indexOf('href="omnia-tokens.css"')).toBeLessThan(html.indexOf('href="taskpane.css"'));
    expect(html.indexOf('href="taskpane.css"')).toBeLessThan(html.indexOf("office.js"));
  });
  it("uses only shared token colors in task pane CSS", () => {
    const css = readFileSync("src/taskpane/taskpane.css", "utf8");
    expect(css).not.toMatch(/#[\da-f]{3,8}\b|rgba?\(|hsla?\(/i);
  });
});
