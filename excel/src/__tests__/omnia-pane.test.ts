import { describe, expect, it } from "vitest";
import { readFileSync, existsSync } from "node:fs";
import { createHash } from "node:crypto";
import { createRequire } from "node:module";
const { JSDOM } = createRequire(import.meta.url)("jsdom");
const read = (name: string) => readFileSync(`src/taskpane/${name}`, "utf8");
const html = read("taskpane.html");
const pane = new JSDOM(html).window.document;
const shell = () => existsSync("src/taskpane/omnia-pane.css") ? read("omnia-pane.css") : "";
describe("byte-identical Omnia pane visual contract", () => {
  it("renders the same icon-free privacy label as the other apps", () => {
    const source = read("taskpane.ts");
    expect(source).toContain("badge.textContent = info.text;");
  });
  it("pins the shared pane and immutable token hashes and load order", () => {
    expect(createHash("sha256").update(shell()).digest("hex")).toBe("643c460b1dc67bd026b7ea7b5e783e4b03dafae298e71745c3aeaef93fe7445b");
    expect(createHash("sha256").update(readFileSync("src/taskpane/omnia-tokens.css")).digest("hex")).toBe("9d4a6a258b587d6c65c9773521ef84e3442915027e052d1f2e82725242cfc3d5");
    expect(Array.from(pane.querySelectorAll('link[rel="stylesheet"]')).map((el: any) => el.getAttribute("href"))).toEqual(["omnia-tokens.css", "omnia-pane.css", "taskpane.css"]);
  });
  it("keeps the shared header skeleton and inline gear without the removed profile chip", () => {
    const header = pane.querySelector(".topbar");
    const structure = (el: any): any => [el.tagName.toLowerCase(), el.className?.baseVal ?? el.className, Array.from(el.children).filter((child: any) => child.tagName.toLowerCase() !== "svg").map(structure)];
    expect(structure(header)).toEqual(["header", "topbar", [
      ["div", "header-main", [["img", "brand-mark", []], ["h1", "", []], ["button", "icon-button", []]]],
      ["div", "header-status", [["p", "", []], ["span", "privacy-badge privacy-local", []]]]
    ]]);
    expect(header.querySelector("h1").textContent).toBe("Sidecar for Excel");
    expect(header.querySelector(".brand-mark").getAttribute("width")).toBe("24");
    expect(header.querySelector("#settings-button").innerHTML).toBe(new JSDOM('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m9.5 4 .7-2h3.6l.7 2 2 .9 2.1-.4 1.8 3.1-1.4 1.6.2 2.3 1.4 1.6-1.8 3.1-2.1-.4-2 .9-.7 2h-3.6l-.7-2-2-.9-2.1.4-1.8-3.1L5 13.5l-.2-2.3-1.4-1.6 1.8-3.1 2.1.4z"/><circle cx="12" cy="11" r="3"/></svg>').window.document.body.innerHTML);
    expect(header.querySelector(".header-status #connection-status")).not.toBeNull();
    expect(header.querySelector(".header-status #privacy-badge")).not.toBeNull();
    for (const id of ["settings-button", "connection-status", "privacy-badge", "bridge-status", "bridge-session-id", "copy-bridge-session", "messages", "prompt-input", "stop-button", "send-button", "settings-dialog", "save-settings-button"]) expect(pane.querySelectorAll(`[id="${id}"]`).length).toBe(1);
  });
  it("uses a three-line autosizing composer without a grip and a visibly disabled Stop", () => {
    expect(pane.querySelector("#prompt-input").getAttribute("rows")).toBe("3");
    expect(shell()).toMatch(/textarea\s*\{[^}]*resize:\s*none/);
    expect(shell()).toMatch(/textarea\s*\{[^}]*field-sizing:\s*content/);
    expect(shell()).toMatch(/\.secondary-button:disabled\s*\{[^}]*cursor:\s*not-allowed[^}]*color:\s*var\(--dim\)/);
    expect(shell()).not.toMatch(/margin-top:\s*auto/);
    expect(pane.querySelector("#stop-button").hasAttribute("disabled")).toBe(true);
  });
  it("stacks settings choices at 200-percent-equivalent narrow widths", () => {
    expect(shell()).toMatch(/@media\s*\(max-width:\s*260px\)\s*\{[\s\S]*grid-template-columns:\s*1fr/);
    expect(shell()).toMatch(/\.appearance, \.appearance-options\s*\{[^}]*flex-direction:\s*column/);
  });
  it("scrolls settings content independently with a pinned Save and no literal colors", () => {
    expect(pane.querySelector(".settings-content")).not.toBeNull();
    expect(pane.querySelector(".settings-panel > menu #save-settings-button")).not.toBeNull();
    expect(shell()).toMatch(/\.settings-content\s*\{[^}]*overflow-y:\s*auto/);
    expect(shell()).toMatch(/\.icon-button\s*\{[^}]*width:\s*36px[^}]*height:\s*36px/);
    expect(shell() + read("taskpane.css")).not.toMatch(/#[\da-f]{3,8}\b|\brgba?\(|\bhsla?\(/i);
  });
});
