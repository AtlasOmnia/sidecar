import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, it } from "vitest";

it("keeps installation and live-acceptance claims bounded across public docs", () => {
  const root = resolve(__dirname, "../..");
  const read = (path: string) => readFileSync(resolve(root, path), "utf8");
  const readme = read("README.md");
  expect(read("plugins/atlas-sidecar-excel/README.md")).toBe(readme);
  for (const path of ["README.md", "docs/desk-bridge.md", "AGENTS.md", "catalog/atlas-sidecar-excel.yaml"]) {
    const text = read(path);
    expect(text, path).toContain("Windows desktop Excel is untested");
    expect(text.toLowerCase(), path).toContain("reopen the sidebar");
    expect(text, path).toContain("no real undo");
  }
  expect(readme).not.toContain("nothing changes without your approval");
  expect(readme).not.toContain("a live run in desktop Excel with this build has not been recorded");
  expect(readme).toContain("Keychain Access > login > Certificates");
  expect(readme).toContain("System keychain");
  expect(readme).toContain("hermes config set plugins.entries.atlas-sidecar-excel.settings.addin_base https://sidebar.example.org/excel/");
  expect(readme).toContain("https://atlasomnia.github.io/sidecar-excel/manifest.xml");
  expect(readme).toContain("the public GitHub Pages origin is first exercised at publication");
  expect(read("plugins/atlas-sidecar-excel/desk_app.py")).toContain('DEFAULT_ADDIN_BASE = "https://atlasomnia.github.io/sidecar-excel/"');
  expect(read("plugins/atlas-sidecar-excel/plugin.yaml")).toContain('default: "https://atlasomnia.github.io/sidecar-excel/"');
  expect(read(".github/workflows/pages.yml")).toContain("npm ci && npm run build:static -- --base https://atlasomnia.github.io/sidecar-excel/");
  expect(read("LICENSE")).toContain("Apache License");
  expect(read("NOTICE")).toContain("Atlas Omnia");
});
