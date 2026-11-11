// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { describe, expect, it } from "vitest";

// Pin historical pre-theme fixture bytes; fresh-root publication has no private Git history.
// Original content hashes and provenance are in fixtures/pre-theme/provenance.json.
const base = (file: string): string => readFileSync(resolve(__dirname, "fixtures/pre-theme", file === "manifest.xml" ? "manifest.xml" : "taskpane.html"), "utf8");
const source = (file: string): string => readFileSync(resolve(__dirname, "../..", file), "utf8");
const oldPane = new DOMParser().parseFromString(base("src/taskpane/taskpane.html"), "text/html");
const pane = new DOMParser().parseFromString(source("src/taskpane/taskpane.html"), "text/html");
const oldManifest = new DOMParser().parseFromString(base("manifest.xml"), "application/xml");
const manifest = new DOMParser().parseFromString(source("manifest.template.xml"), "application/xml");
const elements = (doc: Document, tag: string): Element[] => Array.from(doc.getElementsByTagNameNS("*", tag));

describe("visual-only Omnia contract preservation", () => {
  it("retains all existing element ids and every ARIA, role, type, form and functional attribute", () => {
    const ids = Array.from(oldPane.querySelectorAll<HTMLElement>("[id]"));
    for (const original of ids) {
      if (["base-url-input", "api-key-input"].includes(original.id)) continue;
      const current = pane.getElementById(original.id);
      expect(current, original.id).not.toBeNull();
      for (const attr of Array.from(original.attributes)) {
        if (/^(aria-|role$|type$|rows$|method$|value$|disabled$|hidden$|min$|max$|step$|placeholder$|autocomplete$|spellcheck$)/.test(attr.name)) {
          // Approved polish changes only the prompt's visual default from five lines to three.
          const expected = original.id === "prompt-input" && attr.name === "rows" ? "3" : attr.value;
          expect(current!.getAttribute(attr.name), `${original.id}:${attr.name}`).toBe(expected);
        }
      }
    }
    expect(new Set(Array.from(pane.querySelectorAll("[id]")).map(el => el.id)).size).toBe(pane.querySelectorAll("[id]").length);
    expect(pane.querySelector("form")!.getAttribute("method")).toBe("dialog");
  });
  it("keeps code-dependent existing button, status, label and prompt copy", () => {
    for (const id of ["connection-status", "privacy-badge", "bridge-status", "copy-bridge-session", "checkpoint-stop", "checkpoint-continue", "pending-plan-stop", "pending-plan-continue", "stop-button", "send-button", "save-settings-button"]) {
      expect(pane.getElementById(id)!.textContent).toBe(oldPane.getElementById(id)!.textContent);
    }
    for (const id of ["model-input", "max-iterations-input", "workbook-plan-rollout-input"]) {
      expect(pane.getElementById(id)!.parentElement!.textContent?.trim()).toBe(oldPane.getElementById(id)!.parentElement!.textContent?.trim());
    }
  });
  it("keeps manifest identity, version, locations, domains, permissions, hosts and requirement sets", () => {
    for (const tag of ["Id", "Version", "DefaultLocale", "Permissions", "Requirements", "Hosts"]) {
      expect(elements(manifest, tag).map(el => el.outerHTML)).toEqual(elements(oldManifest, tag).map(el => el.outerHTML));
    }
    const urls = (doc: Document) => elements(doc, "Url").map(el => el.outerHTML);
    expect(urls(manifest).every(url => url.includes("{{ADDIN_BASE}}"))).toBe(true);
    const ids = (doc: Document) => Array.from(doc.querySelectorAll("[id], [resid]")).map(el => [el.localName, el.getAttribute("id"), el.getAttribute("resid")]);
    expect(ids(manifest)).toEqual(ids(oldManifest));
  });
  it("brands both ribbon override versions without renaming resource identifiers", () => {
    expect(elements(manifest, "DisplayName")[0].getAttribute("DefaultValue")).toBe("Sidecar for Excel");
    for (const id of ["GetStarted.Title", "Hermes.Group.Label", "Hermes.OpenPane.Label"]) {
      const matches = elements(manifest, "String").filter(el => el.getAttribute("id") === id);
      expect(matches).toHaveLength(2);
      expect(matches.every(el => el.getAttribute("DefaultValue") === "Sidecar for Excel")).toBe(true);
    }
    expect(elements(manifest, "IconUrl")[0].getAttribute("DefaultValue")).toBe("{{ADDIN_BASE}}assets/icon-32.png");
    expect(elements(manifest, "HighResolutionIconUrl")[0].getAttribute("DefaultValue")).toBe("{{ADDIN_BASE}}assets/icon-128.png");
    for (const image of elements(manifest, "Image").filter(el => el.hasAttribute("DefaultValue"))) {
      expect(image.getAttribute("DefaultValue")).toMatch(/^\{\{ADDIN_BASE\}\}assets\/icon-(16|32|80)\.png$/);
    }
  });
});
