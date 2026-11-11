#!/usr/bin/env node
// SPDX-License-Identifier: Apache-2.0
// Copyright 2026 Atlas Omnia
//
// Build the static Sidecar sidebar for one HTTPS site and render its add-in manifest.
//
//   node scripts/build-static.mjs --base https://example.org/office/word/ [--out dist]
//
// Output (--out, default dist/): taskpane.html, commands.html, js/css, assets/, manifest.xml (from
// manifest.template.xml with this --base) and SHA256SUMS. Host the folder at --base unchanged, set the
// Hermes plugin setting addin_base to the same URL, and load manifest.xml into desktop Office.
// This script never uploads or publishes anything.
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { createRequire } from "node:module";
import { mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repo = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

export function checkBase(base) {
  let url;
  try {
    url = new URL(base);
  } catch {
    throw new Error(`--base must be an https URL ending in /: ${base}`);
  }
  if (url.protocol !== "https:" || !url.hostname || url.username || url.password || url.search || url.hash
      || !url.pathname.endsWith("/") || base !== url.href) {
    throw new Error(`--base must be an https URL ending in / (no query, fragment or user): ${base}`);
  }
  return { base: url.href, origin: url.origin };
}

export function renderManifest(template, base) {
  const { base: checked, origin } = checkBase(base);
  const out = template.replaceAll("{{ADDIN_BASE}}", checked).replaceAll("{{ADDIN_ORIGIN}}", origin)
    .replace(/<!-- Sidecar[\s\S]*?-->\n/, "");
  if (out.includes("{{")) throw new Error("manifest template has an unknown placeholder");
  return out;
}

function walk(dir) {
  return readdirSync(dir).flatMap((name) => {
    const full = join(dir, name);
    return statSync(full).isDirectory() ? walk(full) : [full];
  });
}

export function sha256sums(dir) {
  return walk(dir)
    .filter((file) => !file.endsWith("SHA256SUMS"))
    .map((file) => `${createHash("sha256").update(readFileSync(file)).digest("hex")}  ${relative(dir, file).split("\\").join("/")}`)
    .sort((a, b) => a.slice(66).localeCompare(b.slice(66)))
    .join("\n") + "\n";
}

function main(argv) {
  const args = { out: "dist" };
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i] === "--base") args.base = argv[++i];
    else if (argv[i] === "--out") args.out = argv[++i];
    else throw new Error(`unknown argument ${argv[i]}`);
  }
  if (!args.base) throw new Error("--base <https site ending in /> is required");
  const { base } = checkBase(args.base);
  const out = resolve(repo, args.out);
  const webpackCli = require.resolve("webpack-cli/bin/cli.js");
  const built = spawnSync(process.execPath, [webpackCli, "--mode", "production", "--output-path", out],
    { cwd: repo, stdio: "inherit" });
  if (built.status !== 0) throw new Error("webpack build failed");
  mkdirSync(out, { recursive: true });
  const template = readFileSync(join(repo, "manifest.template.xml"), "utf8");
  writeFileSync(join(out, "manifest.xml"), renderManifest(template, base));
  writeFileSync(join(out, "SHA256SUMS"), sha256sums(out));
  console.log(`Built ${relative(repo, out) || "."} for ${base}`);
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    main(process.argv.slice(2));
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exit(1);
  }
}
