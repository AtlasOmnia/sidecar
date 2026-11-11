// Offline manifest and immutable-theme invariants, not the online Microsoft validator.
// Run: node scripts/verify-omnia-theme.mjs <base-ref> [kit-directory]
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { execFileSync } from "node:child_process";
import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
const require = createRequire(import.meta.url);
const { ManifestHandlerXml } = require("office-addin-manifest/lib/manifestHandler/manifestHandlerXml.js");
const { parseStringPromise } = require("xml2js");
const { JSDOM } = require("jsdom");
const base = process.argv[2] || "HEAD";
const kit = process.argv[3] && path.resolve(process.argv[3]);
const name = "Sidecar for PowerPoint";
const expectedTokenHash = "9d4a6a258b587d6c65c9773521ef84e3442915027e052d1f2e82725242cfc3d5";
const sha = bytes => createHash("sha256").update(bytes).digest("hex");
const beforeText = execFileSync("git", ["show", `${base}:manifest.xml`], { encoding: "utf8" });
const before = await parseStringPromise(beforeText);
// Normalize parser-specific null prototypes; XML data has no JS prototype semantics.
const current = structuredClone(await parseStringPromise(await readFile("manifest.xml", "utf8")));
const allowed = structuredClone(before);
allowed.OfficeApp.DisplayName[0].$.DefaultValue = name;
allowed.OfficeApp.HighResolutionIconUrl[0].$.DefaultValue = "https://localhost:3002/assets/icon-128.png";
const resources = allowed.OfficeApp.VersionOverrides[0].Resources[0];
resources["bt:ShortStrings"][0]["bt:String"].forEach(value => { value.$.DefaultValue = name; });
resources["bt:LongStrings"][0]["bt:String"].forEach(value => { value.$.DefaultValue = name; });
assert.deepEqual(current, allowed, "Unexpected manifest mutation: only approved branding may differ");
const parsed = await new ManifestHandlerXml(path.resolve("manifest.xml")).parseManifest();
assert.equal(parsed.displayName, name);
assert.equal(parsed.id, "7d8aa2f9-2ef4-4f35-8ad1-8f1eb8c2f95f");
assert.equal(parsed.version, "1.0.0.0");
assert.equal(parsed.defaultSettings.sourceLocation, "https://localhost:3002/taskpane.html");
assert.equal(parsed.permissions, "ReadWriteDocument");
assert.deepEqual(parsed.hosts, ["Presentation"]);
const ids = new Set();
const refs = [];
function walk(value) {
  if (Array.isArray(value)) return value.forEach(walk);
  if (!value || typeof value !== "object") return;
  if (value.$?.id) { assert(!ids.has(value.$.id), `Duplicate manifest id: ${value.$.id}`); ids.add(value.$.id); }
  if (value.$?.resid) refs.push(value.$.resid);
  for (const [key, item] of Object.entries(value)) if (key !== "$") walk(item);
}
walk(current);
refs.forEach(ref => assert(ids.has(ref), `Unresolved manifest resource: ${ref}`));
const icons = [];
for (const size of [16, 20, 24, 32, 40, 48, 64, 80, 128]) {
  const filename = `icon-${size}.png`;
  const bytes = await readFile(path.join("assets", filename));
  assert.equal(bytes.toString("hex", 0, 8), "89504e470d0a1a0a");
  assert.equal(bytes.readUInt32BE(16), size); assert.equal(bytes.readUInt32BE(20), size);
  assert.deepEqual(await readFile(path.join("dist/assets", filename)), bytes, "Built icon differs");
  if (kit) assert.deepEqual(await readFile(path.join(kit, "icons", filename)), bytes, "Icon differs from kit");
  icons.push({ filename, sha256: sha(bytes) });
}
const token = await readFile("src/taskpane/omnia-tokens.css");
assert.equal(sha(token), expectedTokenHash);
assert.equal(sha(await readFile("dist/omnia-tokens.css")), expectedTokenHash);
if (kit) assert.deepEqual(await readFile(path.join(kit, "omnia-tokens.css")), token);
for (const filename of ["taskpane.css", "theme-bootstrap.js"]) assert.deepEqual(await readFile(`dist/${filename}`), await readFile(`src/taskpane/${filename}`));
const beforeHtml = new JSDOM(execFileSync("git", ["show", `${base}:src/taskpane/taskpane.html`], { encoding: "utf8" })).window.document;
const afterHtml = new JSDOM(await readFile("src/taskpane/taskpane.html", "utf8")).window.document;
const preservedIds = [];
for (const element of beforeHtml.querySelectorAll("[id]")) {
  const matches = afterHtml.querySelectorAll(`[id="${element.id}"]`);
  assert.equal(matches.length, 1, `Missing/duplicate id: ${element.id}`);
  for (const attribute of element.attributes) {
    if (attribute.name.startsWith("aria-")) assert.equal(matches[0].getAttribute(attribute.name), attribute.value, `Changed ${element.id} ${attribute.name}`);
  }
  preservedIds.push(element.id);
}
assert.equal(afterHtml.querySelector(".bridge-strip").getAttribute("aria-live"), beforeHtml.querySelector(".bridge-strip").getAttribute("aria-live"));
console.log(JSON.stringify({ verdict: "PASS (offline structural/invariant checks)", base, manifestParser: "Installed office-addin-manifest ManifestHandlerXml", onlineValidation: "Not run: XML validator POSTs to Microsoft; no bundled offline XSD", preservedIds, resourceReferences: refs.length, icons, tokenSha256: expectedTokenHash }, null, 2));
