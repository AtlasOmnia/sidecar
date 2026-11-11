// Offline structural validation only. Microsoft's XML validation service is NOT called.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { execFileSync } from 'node:child_process';
const require = createRequire(import.meta.url);
const { JSDOM } = require('jsdom');
const { OfficeAddinManifest } = require('office-addin-manifest/lib/manifestOperations');
const manifest = await OfficeAddinManifest.readManifestFile('manifest.xml');
assert.equal(manifest.displayName, "Sidecar for Excel");
assert.equal(manifest.id, '0bb2b56e-c997-4ef7-a67a-c1a0f3ca4f2e');
assert.equal(manifest.version, '1.0.0.0');
assert.equal(manifest.defaultSettings.sourceLocation, 'https://localhost:3000/taskpane.html');
assert.equal(manifest.permissions, 'ReadWriteDocument');
const xml = new JSDOM(fs.readFileSync('manifest.xml', 'utf8'), { contentType: 'application/xml' }).window.document;
const base = new JSDOM(execFileSync('git', ['show', 'c83ee5cf78a8a473d2131643c661e595baaedeca:manifest.xml'], { encoding: 'utf8' }), { contentType: 'application/xml' }).window.document;
const elements = (doc, tag) => [...doc.getElementsByTagNameNS('*', tag)];
for (const tag of ['Id', 'Version', 'ProviderName', 'DefaultLocale', 'SourceLocation', 'SupportUrl', 'AppDomain', 'Permissions', 'Requirements', 'Hosts', 'Url']) {
  assert.deepEqual(elements(xml, tag).map(el => el.outerHTML), elements(base, tag).map(el => el.outerHTML), tag + ' changed');
}
const ids = new Set([...xml.querySelectorAll('[id]')].map(el => el.getAttribute('id')));
for (const el of xml.querySelectorAll('[resid]')) assert.ok(ids.has(el.getAttribute('resid')), 'Missing resource: ' + el.getAttribute('resid'));
const iconUrls = ['IconUrl', 'HighResolutionIconUrl', 'Image'].flatMap(tag => elements(xml, tag).filter(el => el.hasAttribute('DefaultValue')).map(el => el.getAttribute('DefaultValue')));
for (const url of iconUrls) {
  const asset = url.replace('https://localhost:3000/', '');
  assert.notEqual(asset, url, 'Icon host changed');
  assert.ok(fs.existsSync(asset), 'Missing icon ' + asset);
}
const hash = file => createHash('sha256').update(fs.readFileSync(file)).digest('hex');
const tokenHash = hash('src/taskpane/omnia-tokens.css');
assert.equal(tokenHash, '9d4a6a258b587d6c65c9773521ef84e3442915027e052d1f2e82725242cfc3d5');
const icons = [];
for (const size of [16, 20, 24, 32, 40, 48, 64, 80, 128]) {
  const file = `assets/icon-${size}.png`; const data = fs.readFileSync(file);
  assert.equal(data.readUInt32BE(16), size); assert.equal(data.readUInt32BE(20), size);
  if (process.argv[2]) assert.equal(hash(file), hash(path.join(process.argv[2], 'icons', `icon-${size}.png`)));
  icons.push({ file, size, sha256: hash(file) });
}
if (process.argv[2]) assert.equal(tokenHash, hash(path.join(process.argv[2], 'omnia-tokens.css')));
console.log(JSON.stringify({ verdict: 'PASS (offline parser + structural invariants; not Microsoft service validation)', manifest: { id: manifest.id, version: manifest.version, displayName: manifest.displayName, sourceLocation: manifest.defaultSettings.sourceLocation }, iconUrlsChecked: iconUrls.length, tokenHash, icons }, null, 2));
