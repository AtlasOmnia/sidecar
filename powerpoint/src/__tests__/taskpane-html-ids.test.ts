import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

const root = resolve(__dirname, "..", "taskpane");
const source = readFileSync(resolve(root, "taskpane.ts"), "utf8");
const html = readFileSync(resolve(root, "taskpane.html"), "utf8");

describe("task pane markup", () => {
  it("contains every element the task pane script looks up by id", () => {
    const wanted = new Set<string>();
    for (const match of source.matchAll(/byId<[A-Za-z]+>\("([^"]+)"\)/g)) wanted.add(match[1]);
    expect(wanted.size).toBeGreaterThan(5);
    const missing = [...wanted].filter((id) => !html.includes(`id="${id}"`));
    // A missing element makes byId throw inside Office.onReady, leaving the pane at its static text.
    expect(missing).toEqual([]);
  });

  it("has the Connect to Hermes panel between the bridge strip and the messages", () => {
    const bridge = html.indexOf('class="bridge-strip"');
    const connect = html.indexOf('id="desk-connect"');
    const messages = html.indexOf('id="messages"');
    expect(bridge).toBeGreaterThan(-1);
    expect(connect).toBeGreaterThan(bridge);
    expect(messages).toBeGreaterThan(connect);
  });
});
