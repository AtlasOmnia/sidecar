// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { checkBase, renderManifest } from "../../scripts/build-static.mjs";

const BASE = "https://atlasomnia.github.io/sidecar-powerpoint/";
const ORIGIN = "https://atlasomnia.github.io";
const template = readFileSync("manifest.template.xml", "utf8");
const xml = new DOMParser().parseFromString(renderManifest(template, BASE), "application/xml");
const ns = "http://schemas.microsoft.com/office/appforoffice/1.1";
const bt = "http://schemas.microsoft.com/office/officeappbasictypes/1.0";
const name = "Sidecar for PowerPoint";
function element(tag: string): Element {
  const node = xml.getElementsByTagNameNS(ns, tag)[0];
  expect(node, `Missing manifest ${tag}`).toBeDefined();
  return node;
}

describe("Sidecar for PowerPoint manifest template (offline contract, not the online Office validator)", () => {
  it("renders XML with its own add-in identity, version and permission", () => {
    expect(xml.querySelector("parsererror")).toBeNull();
    expect(element("Id").textContent).toBe("7d8aa2f9-2ef4-4f35-8ad1-8f1eb8c2f95f");
    expect(element("Version").textContent).toBe("1.1.0.0");
    expect(element("Permissions").textContent).toBe("ReadWriteDocument");
    expect(element("DisplayName").getAttribute("DefaultValue")).toBe(name);
  });
  it("points every URL at the configured site and nothing else", () => {
    expect(element("SourceLocation").getAttribute("DefaultValue")).toBe(`${BASE}taskpane.html`);
    expect(Array.from(xml.getElementsByTagNameNS(ns, "AppDomain")).map(node => node.textContent)).toEqual([ORIGIN]);
    const set = element("Set");
    expect(set.getAttribute("Name")).toBe("PowerPointApi");
    expect(set.getAttribute("MinVersion")).toBe("1.4");
    const rendered = renderManifest(template, BASE);
    expect(rendered).not.toMatch(/localhost|\{\{|Omnia's|AppSource|hermesagenttools/);
  });
  it("keeps the remote sidebar site separate from verified localhost discovery", () => {
    const discovery = readFileSync("src/deskConnect.ts", "utf8");
    expect(discovery).toContain("const base = `https://localhost:${port}`");
    expect(discovery).not.toContain("http://127.0.0.1:");
    expect(element("SourceLocation").getAttribute("DefaultValue")).toBe(`${BASE}taskpane.html`);
    expect(Array.from(xml.getElementsByTagNameNS(ns, "AppDomain")).map(node => node.textContent)).toEqual([ORIGIN]);
    const plugin = readFileSync("plugins/atlas-sidecar-powerpoint/plugin.yaml", "utf8");
    expect(plugin).toContain("cryptography>=43");
    expect(plugin).toContain(BASE);
  });
  it("resolves every resource reference", () => {
    const resources = new Map(["Image", "Url", "String"].flatMap(tag => Array.from(xml.getElementsByTagNameNS(bt, tag))).map(node => [node.getAttribute("id"), node]));
    for (const node of Array.from(xml.querySelectorAll("[resid]"))) expect(resources.has(node.getAttribute("resid"))).toBe(true);
    for (const node of Array.from(xml.getElementsByTagNameNS(bt, "Url"))) expect(node.getAttribute("DefaultValue")!.startsWith(BASE)).toBe(true);
    for (const id of ["GetStarted.Title", "Hermes.Group.Label", "Hermes.OpenPane.Label"]) expect(resources.get(id)!.getAttribute("DefaultValue")).toBe(name);
  });
  it("uses the shipped icon sizes", () => {
    expect(element("IconUrl").getAttribute("DefaultValue")).toBe(`${BASE}assets/icon-32.png`);
    expect(element("HighResolutionIconUrl").getAttribute("DefaultValue")).toBe(`${BASE}assets/icon-128.png`);
    for (const size of [16, 20, 24, 32, 40, 48, 64, 80, 128]) {
      const png = readFileSync(`assets/icon-${size}.png`);
      expect(png.subarray(0, 8).toString("hex")).toBe("89504e470d0a1a0a");
      expect(png.readUInt32BE(16)).toBe(size);
      expect(png.readUInt32BE(20)).toBe(size);
    }
  });
  it("refuses a base that is not an https site folder", () => {
    for (const bad of ["http://example.org/x/", "https://example.org/x", "https://u:p@example.org/x/", "https://example.org/x/?q=1"]) {
      expect(() => checkBase(bad)).toThrow();
    }
    expect(checkBase(BASE)).toEqual({ base: BASE, origin: ORIGIN });
  });
});
