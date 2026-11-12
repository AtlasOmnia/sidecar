import { readFileSync } from "node:fs";
import { expect, it } from "vitest";
it.each(["README.md", "plugins/atlas-sidecar-powerpoint/README.md", "docs/desk-bridge.md", "catalog/atlas-sidecar-powerpoint.yaml"])("%s states the live-test and operational boundaries", (path) => {
  const text = readFileSync(path, "utf8").replace(/\s+/g, " ");
  expect(text).toContain("Windows desktop PowerPoint is untested");
  expect(text).toContain("reopen the sidebar");
  expect(text).toContain("no real undo");
  expect(text).toContain("Keychain Access > login > Certificates");
  expect(text).toContain("Sidecar local CA");
  expect(text).toContain("addin_base");
  for (const fact of ["3682e5f", "24 seconds", "three bullets", "315f1f0", "Try again reconnected", "needs your OK", "login keychain only", "plugin-served same-origin default in 1.2.0 has not yet been live-tested in Office"])
    expect(text).toContain(fact);
  expect(text).not.toContain("has not been live-retested");
  expect(text).not.toContain("nothing changes without your approval");
  expect(text).not.toContain("Built and tested for desktop PowerPoint (Microsoft 365 on Windows or macOS)");
});
