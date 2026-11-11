// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { createHash } from "node:crypto";
import { afterEach, describe, expect, it, vi } from "vitest";
import { applyTheme, installThemeControls, loadTheme, normalizeTheme, resolveScheme, saveTheme, type OfficeThemeSnapshot } from "../taskpane/omniaTheme";

const html = readFileSync(resolve(__dirname, "../taskpane/taskpane.html"), "utf8");
const defaults = { palette: "teal", mode: "auto" } as const;
afterEach(() => { vi.useRealTimers(); localStorage.clear(); document.documentElement.removeAttribute("data-scheme"); });

describe("Omnia preferences and scheme resolution", () => {
  it("defaults to Teal + Auto without stored preferences", () => {
    expect(loadTheme(localStorage)).toEqual(defaults);
  });
  it.each(["teal", "blue", "gold", "red"])("loads the stored %s palette", palette => {
    localStorage.setItem("omnia.theme", JSON.stringify({ palette, mode: "light" }));
    expect(loadTheme(localStorage)).toEqual({ palette, mode: "light" });
  });
  it.each([null, [], "dark", { palette: "oops", mode: "oops" }, { palette: "toString", mode: "__proto__" }])("rejects malformed preferences %j", value => {
    expect(normalizeTheme(value)).toEqual(defaults);
  });
  it("normalizes fields independently and ignores unrelated stored properties", () => {
    expect(normalizeTheme({ palette: "red", mode: "invalid", key: "ignored" })).toEqual({ palette: "red", mode: "auto" });
  });
  it("survives malformed JSON or denied storage", () => {
    localStorage.setItem("omnia.theme", "{");
    expect(loadTheme(localStorage)).toEqual(defaults);
    expect(loadTheme({ getItem() { throw new Error("denied"); } })).toEqual(defaults);
    expect(saveTheme({ setItem() { throw new Error("denied"); } }, defaults)).toBe(false);
  });
  it("persists only palette and mode to the exact key", () => {
    expect(saveTheme(localStorage, { palette: "gold", mode: "dark" })).toBe(true);
    expect(JSON.parse(localStorage.getItem("omnia.theme")!)).toEqual({ palette: "gold", mode: "dark" });
    expect(localStorage.length).toBe(1);
  });
  it.each([true, false])("Auto respects Office isDarkTheme=%s over background and OS", isDarkTheme => {
    expect(resolveScheme("auto", { isDarkTheme, bodyBackgroundColor: isDarkTheme ? "#ffffff" : "#000000" }, !isDarkTheme)).toBe(isDarkTheme ? "dark" : "light");
  });
  it.each([["#000000", "dark"], ["#FFFFFF", "light"], ["#222222", "dark"], ["#f4f7f9", "light"], ["#fff", "light"], ["#333", "dark"]])("Auto uses Office background %s luminance", (bodyBackgroundColor, expected) => {
    expect(resolveScheme("auto", { bodyBackgroundColor }, expected !== "dark")).toBe(expected);
  });
  it.each([undefined, {}, { bodyBackgroundColor: "not-a-color" }])("falls back to OS when Office theme is unavailable %j", theme => {
    expect(resolveScheme("auto", theme, true)).toBe("dark");
    expect(resolveScheme("auto", theme, false)).toBe("light");
  });
  it.each(["dark", "light"] as const)("explicit %s wins over Office and OS", mode => {
    expect(resolveScheme(mode, { isDarkTheme: mode !== "dark" }, mode !== "dark")).toBe(mode);
  });
  it("applies palette and resolved scheme on html without replacing other attributes", () => {
    document.documentElement.lang = "en";
    applyTheme(document.documentElement, { palette: "blue", mode: "auto" }, { isDarkTheme: false }, true);
    expect(document.documentElement.dataset.palette).toBe("blue");
    expect(document.documentElement.dataset.scheme).toBe("light");
    expect(document.documentElement.lang).toBe("en");
  });
});

describe("native settings controls", () => {
  function setup() {
    document.body.innerHTML = new DOMParser().parseFromString(html, "text/html").body.innerHTML;
    let office: OfficeThemeSnapshot | undefined = { isDarkTheme: true };
    const media = Object.assign(new EventTarget(), { matches: false });
    const controller = installThemeControls({ document, storage: localStorage, media, getOfficeTheme: () => office });
    return { controller, media, setOffice: (value: OfficeThemeSnapshot | undefined) => { office = value; } };
  }
  it("offers four palette swatches and three appearance choices with labelled native radio groups", () => {
    const { controller } = setup();
    expect(Array.from(document.querySelectorAll<HTMLInputElement>('[name="omnia-palette"]')).map(x => x.value)).toEqual(["teal", "blue", "gold", "red"]);
    expect(Array.from(document.querySelectorAll<HTMLInputElement>('[name="omnia-mode"]')).map(x => x.value)).toEqual(["auto", "dark", "light"]);
    expect(document.querySelectorAll("fieldset legend").length).toBe(2);
    expect(document.querySelector<HTMLInputElement>('[name="omnia-palette"][value="teal"]')!.checked).toBe(true);
    expect(document.querySelector<HTMLInputElement>('[name="omnia-mode"][value="auto"]')!.checked).toBe(true);
    controller.dispose();
  });
  it("applies and persists control changes immediately, without submitting the connection settings", () => {
    const { controller } = setup();
    document.querySelector<HTMLInputElement>('[name="omnia-palette"][value="gold"]')!.click();
    document.querySelector<HTMLInputElement>('[name="omnia-mode"][value="light"]')!.click();
    expect(document.documentElement.dataset.palette).toBe("gold");
    expect(document.documentElement.dataset.scheme).toBe("light");
    expect(loadTheme(localStorage)).toEqual({ palette: "gold", mode: "light" });
    expect(localStorage.getItem("hermes.settings")).toBeNull();
    controller.dispose();
  });
  it("reflects stored selection and stays usable when storage writes fail", () => {
    localStorage.setItem("omnia.theme", JSON.stringify({ palette: "red", mode: "dark" }));
    const { controller } = setup();
    expect(document.querySelector<HTMLInputElement>('[name="omnia-palette"][value="red"]')!.checked).toBe(true);
    const denied = vi.spyOn(Storage.prototype, "setItem").mockImplementation(() => { throw new Error("denied"); });
    document.querySelector<HTMLInputElement>('[name="omnia-mode"][value="light"]')!.click();
    expect(document.documentElement.dataset.scheme).toBe("light");
    denied.mockRestore(); controller.dispose();
  });
  it("refreshes Auto as Office theme changes and responds to OS fallback changes; dispose removes polling", () => {
    vi.useFakeTimers();
    const { controller, media, setOffice } = setup();
    expect(document.documentElement.dataset.scheme).toBe("dark");
    setOffice({ isDarkTheme: false });
    vi.advanceTimersByTime(1100);
    expect(document.documentElement.dataset.scheme).toBe("light");
    setOffice(undefined); media.matches = true; media.dispatchEvent(new Event("change"));
    expect(document.documentElement.dataset.scheme).toBe("dark");
    controller.dispose(); media.matches = false; media.dispatchEvent(new Event("change"));
    vi.advanceTimersByTime(1100);
    expect(document.documentElement.dataset.scheme).toBe("dark");
  });
});

describe("immutable shared kit and first paint", () => {
  it("vendors the approved token block byte-for-byte", () => {
    const tokens = readFileSync(resolve(__dirname, "../taskpane/omnia-tokens.css"));
    expect(createHash("sha256").update(tokens).digest("hex")).toBe("9d4a6a258b587d6c65c9773521ef84e3442915027e052d1f2e82725242cfc3d5");
  });
  it("loads tokens before pane CSS and executes theme bootstrap synchronously before Office.js", () => {
    expect(html.indexOf('href="omnia-tokens.css"')).toBeGreaterThan(0);
    expect(html.indexOf('href="omnia-tokens.css"')).toBeLessThan(html.indexOf('href="taskpane.css"'));
    expect(html.indexOf('id="omnia-theme-bootstrap"')).toBeLessThan(html.indexOf("appsforoffice.microsoft.com"));
    expect(html).toMatch(/<script id="omnia-theme-bootstrap" src="[^\n]+"><\/script>/);
  });
  it("uses no literal CSS colors outside the unchanged tokens", () => {
    const css = readFileSync(resolve(__dirname, "../taskpane/taskpane.css"), "utf8");
    expect(css).not.toMatch(/#[\da-f]{3,8}\b|rgba?\(|hsla?\(/i);
  });
});
