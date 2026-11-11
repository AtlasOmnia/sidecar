// SPDX-License-Identifier: Apache-2.0
import { readFileSync } from "node:fs";
import { expect, it } from "vitest";
import { ApprovalRequiredError, DESK_COMMAND } from "../deskConnect";

it("names the actual Sidecar approval command", () => {
  expect(DESK_COMMAND).toBe("/sidecar-powerpoint");
  expect(new ApprovalRequiredError().message).toContain("/sidecar-powerpoint");
});

it("keeps installed documentation identical and discloses TLS/agent authority", () => {
  const readme = readFileSync("README.md", "utf8");
  expect(readFileSync("plugins/atlas-sidecar-powerpoint/README.md", "utf8")).toBe(readme);
  for (const path of ["README.md", "catalog/atlas-sidecar-powerpoint.yaml", "plugins/atlas-sidecar-powerpoint/desk_tls.py"]) {
    const text = readFileSync(path, "utf8");
    expect(text).toContain("For TLS it can only vouch for localhost. The trust command does not limit the certificate to TLS, so protect the key file and remove the trust when you stop using Sidecar.");
  }
  expect(readme).toContain("all its tools until you revoke");
  expect(readme).toContain("plugins.entries.atlas-sidecar-powerpoint.settings.addin_base");
});

it("prepares the decided Pages build without an author-private default", () => {
  const base = "https://atlasomnia.github.io/sidecar-powerpoint/";
  for (const path of ["plugins/atlas-sidecar-powerpoint/desk_app.py", "plugins/atlas-sidecar-powerpoint/plugin.yaml"])
    expect(readFileSync(path, "utf8")).toContain(base);
  expect(readFileSync(".github/workflows/pages.yml", "utf8")).toContain(`npm ci && npm run build:static -- --base ${base}`);
  const manifest = readFileSync("manifest.template.xml", "utf8");
  expect(manifest).toContain('id="GetStarted.Description" DefaultValue="Connect your Hermes agent to read and edit this presentation from the sidebar."');
  expect(manifest).toContain('id="Hermes.OpenPane.Tooltip" DefaultValue="Open the sidebar to chat with your Hermes agent about this presentation."');
});
