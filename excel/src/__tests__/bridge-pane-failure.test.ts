import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import vm from "node:vm";
import ts from "typescript";
import { describe, expect, it } from "vitest";

const source = readFileSync(resolve(__dirname, "../taskpane/taskpane.ts"), "utf8");
// Run the production functions, isolating Office startup and unrelated workbook UI.
const ast = ts.createSourceFile("taskpane.ts", source, ts.ScriptTarget.Latest, true);
const functions = ast.statements.filter((node) => ts.isFunctionDeclaration(node)
  && ["setConnectPanel", "updateBridgeState", "bindControls"].includes(node.name?.text ?? ""))
  .map((node) => node.getText(ast)).join("\n");
const code = ts.transpileModule(functions, { compilerOptions: { target: ts.ScriptTarget.ES2022 } }).outputText;
function pane(grant: string | null = "test-only-grant") {
  const elements = new Map<string, any>();
  const byId = (id: string) => {
    if (!elements.has(id)) elements.set(id, { hidden: false, disabled: false, textContent: "", removeAttribute() {}, addEventListener(event: string, handler: () => void) { this[event] = handler; } });
    return elements.get(id);
  };
  let retries = 0;
  const context = vm.createContext({
    connectToHermes: async () => { retries += 1; },
    byId, connector: { grant }, bridgeBase: "https://localhost:47640", DESK_COMMAND: "/sidecar-excel",
    setStatus() {}, turnCoordinator: { bridgeStateInvalidatesActiveTurn: () => false },
    activeTurnOperationId: null, lastBridgeSessionId: null, lastBridgeDocumentId: null,
  });
  vm.runInContext(code, context);
  vm.runInContext("bindControls()", context);
  return { elements, retry() { elements.get("desk-connect-button").click(); return retries; }, update(state: object) { context.state = state; vm.runInContext("updateBridgeState(state)", context); } };
}

describe("bridge failure panel", () => {
  it("does not claim lost approval for an approved pane's transport failure", () => {
    const p = pane();
    p.update({ connected: false, sessionId: null, documentId: null, lastError: "private-error-test" });
    expect(p.elements.get("desk-connect-text").textContent).toBe("Lost the connection to Hermes on this computer.");
    expect(p.elements.get("desk-connect-button").textContent).toBe("Try again");
    expect(p.elements.get("desk-connect-text").textContent).not.toContain("private-error-test");
    expect(p.retry()).toBe(1);
  });
  it("shows approval copy only without a grant, and hides the panel on connection", () => {
    const p = pane(null);
    p.update({ connected: false, sessionId: null, documentId: null, lastError: "approval required" });
    expect(p.elements.get("desk-connect-text").textContent).toBe("This sidebar needs your OK in Hermes before it can chat.");
    p.update({ connected: true, sessionId: "session", documentId: "document" });
    expect(p.elements.get("desk-connect").hidden).toBe(true);
  });
});
