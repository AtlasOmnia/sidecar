"use strict";

const { spawnSync } = require("child_process");
const { win32 } = require("path");

function resolveBash({ platform = process.platform, env = process.env, spawn = spawnSync } = {}) {
  let candidates = ["bash"];
  if (platform === "win32") {
    // CreateProcess can choose System32's WSL launcher even when Git Bash is
    // on PATH. Only probe absolute paths derived from Git installations.
    candidates = [];
    for (const base of [env.ProgramFiles, env["ProgramFiles(x86)"], env.LOCALAPPDATA && win32.join(env.LOCALAPPDATA, "Programs")]) {
      if (base) candidates.push(win32.join(base, "Git", "bin", "bash.exe"));
    }
    const gitPaths = spawn("where.exe", ["git.exe"], { encoding: "utf8", shell: false, timeout: 5000, env });
    if (!gitPaths.error && gitPaths.status === 0) {
      for (const git of (gitPaths.stdout ?? "").trim().split(/\r?\n/)) {
        if (!win32.isAbsolute(git)) continue;
        const directory = win32.dirname(git);
        candidates.push(win32.resolve(directory, "../bin/bash.exe"));
        candidates.push(win32.resolve(directory, "../../bin/bash.exe"));
      }
    }
  }
  for (const command of new Set(candidates)) {
    const probe = spawn(command, ["--version"], { encoding: "utf8", shell: false, timeout: 5000, env });
    if (!probe.error && probe.status === 0 && probe.stdout?.includes("GNU bash")) return command;
  }
  throw new Error(platform === "win32"
    ? "Git Bash is required to validate shell scripts on Windows; install Git for Windows."
    : "bash is required to validate shell scripts.");
}

module.exports = { resolveBash };
