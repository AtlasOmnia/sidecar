import { describe, expect, it, vi } from "vitest";
import { createRequire } from "node:module";
// JSDOM is already a dev dependency; its package does not ship TS declarations.
const { JSDOM } = createRequire(import.meta.url)("jsdom");
import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import path from "node:path";

const read = (file: string) => readFileSync(path.resolve(file), "utf8");
function pane(stored?: string, officeTheme?: unknown, systemDark = false) {
  const dom = new JSDOM(read("src/taskpane/taskpane.html"), {
    url: "https://localhost:3002/taskpane.html", runScripts: "outside-only"
  });
  const win = dom.window;
  const media = { matches: systemDark, addEventListener: vi.fn(), removeEventListener: vi.fn() };
  Object.defineProperty(win, "matchMedia", { value: () => media });
  if (stored !== undefined) win.localStorage.setItem("omnia.theme", stored);
  if (officeTheme !== undefined) (win as any).Office = { context: { officeTheme } };
  win.eval(read("src/taskpane/theme-bootstrap.js"));
  win.document.dispatchEvent(new win.Event("DOMContentLoaded"));
  return { dom, win, api: (win as any).OmniaTheme, media, root: win.document.documentElement };
}

describe("Omnia theme production bootstrap", () => {
  it("defaults to Teal + Auto before first paint", () => {
    const { api, root } = pane();
    expect(api.getSettings()).toEqual({ palette: "teal", mode: "auto" });
    expect(root.dataset.palette).toBe("teal");
    expect(root.dataset.scheme).toBe("light");
  });
  it.each(["teal", "blue", "gold", "red"])("restores stored %s + Dark", (palette) => {
    const { api, root } = pane(JSON.stringify({ palette, mode: "dark" }), { isDarkTheme: false });
    expect(api.getSettings()).toEqual({ palette, mode: "dark" });
    expect(root.dataset.palette).toBe(palette);
    expect(root.dataset.scheme).toBe("dark");
  });
  it.each(["{broken", "null", "42", '{"palette":"purple","mode":"sepia"}'])("handles invalid storage %s", (stored) => {
    expect(pane(stored).api.getSettings()).toEqual({ palette: "teal", mode: "auto" });
  });
  it("normalizes stored fields independently", () => {
    expect(pane('{"palette":"gold","mode":"bad"}').api.getSettings()).toEqual({ palette: "gold", mode: "auto" });
  });
  it.each([true, false])("Auto follows Office isDarkTheme=%s over media/color", (dark) => {
    expect(pane(undefined, { isDarkTheme: dark, bodyBackgroundColor: dark ? "#ffffff" : "#000000" }, !dark).root.dataset.scheme).toBe(dark ? "dark" : "light");
  });
  it.each([["#202020", "dark"], ["#ffffff", "light"], ["000000", "dark"], ["#fff", "light"]])("Auto resolves Office body luminance %s", (color, scheme) => {
    expect(pane(undefined, { bodyBackgroundColor: color }, scheme !== "dark").root.dataset.scheme).toBe(scheme);
  });
  it.each([true, false])("Auto falls back to media when Office is unavailable (%s)", (dark) => {
    expect(pane(undefined, undefined, dark).root.dataset.scheme).toBe(dark ? "dark" : "light");
    expect(pane(undefined, { bodyBackgroundColor: "invalid" }, dark).root.dataset.scheme).toBe(dark ? "dark" : "light");
  });
  it("re-resolves after Office readiness and theme changes", () => {
    const { win, api, root } = pane();
    (win as any).Office = { context: { officeTheme: { isDarkTheme: true } } };
    api.refresh();
    expect(root.dataset.scheme).toBe("dark");
    (win as any).Office.context.officeTheme.isDarkTheme = false;
    api.refresh();
    expect(root.dataset.scheme).toBe("light");
  });
  it("responds to media changes only when no Office theme wins", () => {
    const { root, media } = pane();
    expect(media.addEventListener).toHaveBeenCalledWith("change", expect.any(Function));
    media.matches = true;
    media.addEventListener.mock.calls[0][1]();
    expect(root.dataset.scheme).toBe("dark");
  });
  it("persists controls immediately without submitting connection settings", () => {
    const { win, root } = pane();
    const submit = vi.fn();
    win.document.querySelector("form")!.addEventListener("submit", submit);
    (win.document.querySelector('[data-set-palette="gold"]') as HTMLElement).click();
    (win.document.querySelector('[data-set-mode="dark"]') as HTMLElement).click();
    expect(JSON.parse(win.localStorage.getItem("omnia.theme")!)).toEqual({ palette: "gold", mode: "dark" });
    expect(root.dataset.palette).toBe("gold");
    expect(root.dataset.scheme).toBe("dark");
    expect(submit).not.toHaveBeenCalled();
    expect(win.document.querySelector('[data-set-palette="gold"]')!.getAttribute("aria-checked")).toBe("true");
    expect(win.document.querySelector('[data-set-mode="auto"]')!.getAttribute("aria-checked")).toBe("false");
  });
  it("supports roving radio keyboard navigation and wraparound", () => {
    const { win, root } = pane();
    const red = win.document.querySelector('[data-set-palette="red"]') as HTMLElement;
    red.dispatchEvent(new win.KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    expect(root.dataset.palette).toBe("teal");
    const teal = win.document.querySelector('[data-set-palette="teal"]') as HTMLElement;
    expect(win.document.activeElement).toBe(teal);
    teal.dispatchEvent(new win.KeyboardEvent("keydown", { key: "End", bubbles: true }));
    expect(root.dataset.palette).toBe("red");
    expect(red.tabIndex).toBe(0);
    expect(teal.tabIndex).toBe(-1);
  });
  it("applies Light explicitly even with dark Office/media", () => {
    const { root } = pane('{"palette":"blue","mode":"light"}', { isDarkTheme: true }, true);
    expect(root.dataset.scheme).toBe("light");
  });
  it("keeps working when storage read/write is denied", () => {
    const { win, api, root } = pane();
    vi.spyOn(win.Storage.prototype, "setItem").mockImplementation(() => { throw new Error("denied"); });
    expect(() => api.setSettings({ palette: "red", mode: "dark" })).not.toThrow();
    expect(root.dataset.palette).toBe("red");
    vi.spyOn(win.Storage.prototype, "getItem").mockImplementation(() => { throw new Error("denied"); });
    expect(() => win.eval(read("src/taskpane/theme-bootstrap.js"))).not.toThrow();
  });
});

describe("Omnia visual contracts", () => {
  it("vendors the exact approved rev-3 token bytes", () => {
    const bytes = readFileSync("src/taskpane/omnia-tokens.css");
    expect(createHash("sha256").update(bytes).digest("hex")).toBe("9d4a6a258b587d6c65c9773521ef84e3442915027e052d1f2e82725242cfc3d5");
  });
  it("loads bootstrap and tokens before taskpane styling, with no literal colors", () => {
    const html = read("src/taskpane/taskpane.html");
    expect(html.indexOf('src="theme-bootstrap.js"')).toBeLessThan(html.indexOf('href="omnia-tokens.css"'));
    expect(html.indexOf('href="omnia-tokens.css"')).toBeLessThan(html.indexOf('href="taskpane.css"'));
    expect(html).toContain("Sidecar for PowerPoint");
    expect(html).toContain('src="assets/icon-32.png"');
    expect(read("src/taskpane/taskpane.css")).not.toMatch(/#[\da-f]{3,8}\b|\brgba?\(|\bhsla?\(/i);
  });
  it("keeps all previous IDs and ARIA attributes intact", () => {
    const { window } = new JSDOM(read("src/taskpane/taskpane.html"));
    for (const id of ["connection-status", "privacy-badge", "settings-button", "bridge-status", "bridge-session-id", "copy-bridge-session", "messages", "checkpoint-bar", "checkpoint-stop", "checkpoint-continue", "prompt-input", "stop-button", "send-button", "settings-dialog", "close-settings-button", "bridge-url", "model-input", "max-iterations-input", "save-settings-button"]) {
      expect(window.document.querySelectorAll(`[id="${id}"]`).length).toBe(1);
    }
    expect(window.document.querySelector("#messages")!.getAttribute("aria-live")).toBe("polite");
    expect(window.document.querySelector(".bridge-strip")!.getAttribute("aria-live")).toBe("polite");
    expect(window.document.querySelector("#settings-button")!.getAttribute("aria-label")).toBe("Settings");
    expect(window.document.querySelector("#close-settings-button")!.getAttribute("aria-label")).toBe("Close");
  });
});
