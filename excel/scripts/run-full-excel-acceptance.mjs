#!/usr/bin/env node

import { spawn } from "node:child_process";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { requestText, buildRunnerEnv } from "./request-text.mjs";

const ACCEPTANCE_URL = "https://localhost:3000/taskpane.html?fullFunctionalityAcceptance=1";
const READINESS_TIMEOUT_MS = 60_000;
const READINESS_INTERVAL_MS = 250;
const ACCEPTANCE_CHUNK_ID = "src_taskpane_fullFunctionalityAcceptanceHarness_ts";
const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

if (process.argv.includes("--help") || process.argv.includes("-h")) {
  console.log("Run the acceptance-enabled HTTPS task-pane server and print the native harness invocation.");
  console.log("Usage: node scripts/run-full-excel-acceptance.mjs");
  process.exit(0);
}

const packageJson = JSON.parse(
  readFileSync(path.join(REPO_ROOT, "package.json"), "utf8"),
);
if (typeof packageJson.scripts?.["dev-server"] !== "string") {
  throw new Error(`repository root is missing the dev-server script: ${REPO_ROOT}`);
}

const webpackCliPath = require.resolve("webpack-cli/bin/cli.js", {
  paths: [REPO_ROOT],
});
const child = spawn(process.execPath, [webpackCliPath, "serve", "--mode", "development"], {
  cwd: REPO_ROOT,
  env: buildRunnerEnv(),
  stdio: "inherit",
});

let stopping = false;
let childExited = false;
let childFailure = null;
let forceKillTimer = null;

const childExit = new Promise((resolve) => {
  child.once("error", (error) => {
    if (forceKillTimer) clearTimeout(forceKillTimer);
    childFailure = error;
    childExited = true;
    if (!stopping) {
      console.error(error.message);
      process.exitCode = 1;
    }
    resolve();
  });
  child.once("exit", (code, signal) => {
    if (forceKillTimer) clearTimeout(forceKillTimer);
    childExited = true;
    if (!stopping && code !== 0) {
      process.exitCode = code ?? 1;
    }
    resolve({ code, signal });
  });
});

function terminateChild() {
  if (stopping) {
    return;
  }
  stopping = true;
  if (!childExited && child.exitCode === null) {
    child.kill("SIGTERM");
    forceKillTimer = setTimeout(() => {
      if (!childExited && child.exitCode === null) {
        child.kill("SIGKILL");
      }
    }, 5_000);
    forceKillTimer.unref();
  }
}

process.once("SIGINT", terminateChild);
process.once("SIGTERM", terminateChild);

function assertChildRunning() {
  if (childFailure) {
    throw childFailure;
  }
  if (childExited || child.exitCode !== null) {
    throw new Error("acceptance dev server exited before readiness");
  }
}

async function checkReadiness() {
  const html = await requestText(ACCEPTANCE_URL);
  if (html === null) return false;

  const scriptSources = [...html.matchAll(/<script[^>]+src=["']([^"']+)["']/gi)]
    .map((match) => new URL(match[1], ACCEPTANCE_URL))
    .filter((url) => url.origin === "https://localhost:3000");
  for (const scriptUrl of scriptSources) {
    const entrySource = await requestText(scriptUrl);
    if (entrySource === null) continue;
    const chunkHash = entrySource.match(
      new RegExp(`"${ACCEPTANCE_CHUNK_ID}":"([a-f0-9]+)"`),
    )?.[1];
    if (!chunkHash) continue;
    const chunkUrl = new URL(
      `${ACCEPTANCE_CHUNK_ID}.${chunkHash}.js`,
      scriptUrl,
    );
    const chunkSource = await requestText(chunkUrl);
    if (
      chunkSource?.includes("__hermesExcelFullFunctionalityAcceptance") &&
      chunkSource.includes("Hermes_Acceptance_")
    ) {
      return true;
    }
  }
  return false;
}

async function waitForExactTaskPaneUrl() {
  const startedAt = Date.now();
  while (Date.now() - startedAt < READINESS_TIMEOUT_MS) {
    if (stopping) {
      throw new Error("acceptance server was stopped");
    }
    assertChildRunning();
    if (await checkReadiness()) {
      // Require both acceptance chunk markers and the directly owned server.
      await new Promise((resolve) => setTimeout(resolve, 500));
      assertChildRunning();
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, READINESS_INTERVAL_MS));
  }
  throw new Error(`timed out waiting for ${ACCEPTANCE_URL}`);
}

try {
  await waitForExactTaskPaneUrl();
  console.log(ACCEPTANCE_URL);
  console.log("await window.__hermesExcelFullFunctionalityAcceptance.run()");
} catch (error) {
  if (!stopping) {
    console.error(error instanceof Error ? error.message : String(error));
  }
  terminateChild();
  process.exitCode = 1;
}

await childExit;
