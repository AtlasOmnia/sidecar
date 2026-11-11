// Isolated built-pane visual harness. No Office, registry, live services, or external network.
// Run: node scripts/capture-omnia-theme.mjs <output-directory> [before-dist-directory]
import { createServer } from "node:http";
import { spawn } from "node:child_process";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import path from "node:path";
import assert from "node:assert/strict";

const repo = process.cwd();
const output = path.resolve(process.argv[2] || ".tmp/captures");
const baseline = path.resolve(process.argv[3] || ".tmp/before-dist");
const profile = path.join(repo, ".tmp", `edge-omnia-${process.pid}`);
const edgePath = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe";
await mkdir(output, { recursive: true });
await mkdir(profile, { recursive: true });
const stub = `
window.fetch = async () => { throw new Error('Visual fixture: network disabled'); };
window.Office = {
  context: { officeTheme: { isDarkTheme: new URLSearchParams(location.search).get('office') !== 'light' } },
  onReady: function (callback) { return Promise.resolve().then(callback); },
  VisibilityMode: { hidden: 'Hidden' }
};
`;
const server = createServer(async (request, response) => {
  try {
    const url = new URL(request.url, "http://127.0.0.1");
    const before = url.pathname.startsWith("/before/");
    const base = before ? baseline : path.join(repo, "dist");
    const name = decodeURIComponent(url.pathname.replace(/^\/(before|after)\//, ""));
    if (name === "office-stub.js") {
      response.setHeader("Content-Type", "text/javascript"); response.end(stub); return;
    }
    const file = path.resolve(base, name);
    if (!file.startsWith(base + path.sep)) throw new Error("Outside fixture root");
    let bytes = await readFile(file);
    if (name.endsWith(".html")) bytes = Buffer.from(bytes.toString().replace("https://appsforoffice.microsoft.com/lib/1/hosted/office.js", "office-stub.js"));
    response.setHeader("Content-Type", name.endsWith(".html") ? "text/html" : name.endsWith(".js") ? "text/javascript" : name.endsWith(".css") ? "text/css" : "image/png");
    response.end(bytes);
  } catch { response.writeHead(404); response.end(); }
});
await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
const port = server.address().port;
assert(port >= 47000, "Fixture must use a free high port");
const edge = spawn(edgePath, ["--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check", "--disable-background-networking", "--disable-component-update", "--disable-sync", "--no-proxy-server", "--remote-debugging-address=127.0.0.1", "--remote-debugging-port=0", `--user-data-dir=${profile}`, "--window-size=380,760", "about:blank"], { stdio: "ignore" });
let ws;
let counter = 0;
const pending = new Map();
const exceptions = [];
const evidence = { serverPort: port, browser: null, captures: [], layoutChecks: [], interactionChecks: [], exceptions };
function delay(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }
async function until(callback, label, timeout = 15000) {
  const end = Date.now() + timeout;
  while (Date.now() < end) { const value = await callback(); if (value) return value; await delay(100); }
  throw new Error(`Timed out: ${label}`);
}
function send(method, params = {}) {
  const id = ++counter;
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => { pending.delete(id); reject(new Error(`CDP timeout: ${method}`)); }, 15000);
    pending.set(id, { resolve, reject, timer });
    ws.send(JSON.stringify({ id, method, params }));
  });
}
async function evaluate(expression) {
  const result = await send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
  if (result.exceptionDetails) throw new Error(JSON.stringify(result.exceptionDetails));
  return result.result.value;
}
const sample = `(() => {
  const messages = document.getElementById('messages');
  messages.innerHTML = '<article class="message user"><div class="bubble">Summarize this deck for the leadership team.</div></article>' +
    '<article class="message assistant"><div class="bubble"><p>Here’s a concise <strong>executive summary</strong> for your presentation.</p><ul><li><strong>Focus:</strong> a clearer story from opportunity to action.</li><li><strong>Progress:</strong> key milestones and the next decision.</li></ul><p>I can also turn this into a closing slide, keeping your existing layout.</p></div><details class="activity" open><summary>Activity</summary><div class="activity-list"><div class="activity-item"><span>✓</span><span>Read presentation overview · complete</span></div></div></details></article>';
  document.getElementById('bridge-status').textContent = 'Office bridge connected';
  document.getElementById('bridge-status').className = 'bridge-status bridge-connected';
  document.getElementById('bridge-session-id').textContent = 'visual-fixture-session';
  document.getElementById('copy-bridge-session').disabled = false;
  document.getElementById('connection-status').textContent = 'Ready';
  messages.scrollTop = 0;
})()`;
async function metrics() {
  return evaluate(`(() => {
    const rect = el => { const r=el.getBoundingClientRect(); return {x:r.x,y:r.y,width:r.width,height:r.height,right:r.right,bottom:r.bottom}; };
    return {width:innerWidth,height:innerHeight,palette:document.documentElement.dataset.palette,scheme:document.documentElement.dataset.scheme,bodyScroll:document.body.scrollWidth,rootScroll:document.documentElement.scrollWidth,composer:rect(document.querySelector('.composer')),send:rect(document.getElementById('send-button')),messages:rect(document.getElementById('messages')),background:getComputedStyle(document.body).backgroundColor};
  })()`);
}
async function navigate(before, palette = "teal", mode = "dark", office = "dark", width = 380) {
  await send("Emulation.setDeviceMetricsOverride", { width, height: 760, deviceScaleFactor: 1, mobile: false });
  await send("Page.navigate", { url: `http://127.0.0.1:${port}/${before ? "before" : "after"}/taskpane.html?palette=${palette}&mode=${mode}&office=${office}` });
  await until(() => evaluate("document.readyState === 'complete' && document.getElementById('connection-status')?.textContent === 'Ready'"), "built pane ready");
  await evaluate(sample);
  await evaluate("new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))");
}
async function capture(filename) {
  const screenshot = await send("Page.captureScreenshot", { format: "png", captureBeyondViewport: false });
  await writeFile(path.join(output, filename), Buffer.from(screenshot.data, "base64"));
  evidence.captures.push({ filename, ...(await metrics()) });
}
try {
  const active = await until(async () => { try { return (await readFile(path.join(profile, "DevToolsActivePort"), "utf8")).split(/\r?\n/); } catch { return null; } }, "Edge CDP");
  assert(Number(active[0]) >= 47000, "Debug endpoint must be high-port");
  const version = await (await fetch(`http://127.0.0.1:${active[0]}/json/version`)).json();
  evidence.browser = version.Browser;
  const targets = await (await fetch(`http://127.0.0.1:${active[0]}/json/list`)).json();
  ws = new WebSocket(targets.find(target => target.type === "page").webSocketDebuggerUrl);
  await new Promise((resolve, reject) => { ws.addEventListener("open", resolve, { once: true }); ws.addEventListener("error", reject, { once: true }); });
  ws.addEventListener("message", event => {
    const message = JSON.parse(event.data);
    if (message.method === "Runtime.exceptionThrown") exceptions.push(message.params.exceptionDetails);
    if (!message.id) return;
    const call = pending.get(message.id);
    if (!call) return;
    clearTimeout(call.timer); pending.delete(message.id);
    if (message.error) call.reject(new Error(JSON.stringify(message.error))); else call.resolve(message.result);
  });
  await send("Page.enable"); await send("Runtime.enable");
  await send("Network.enable");
  await send("Network.setBlockedURLs", { urls: ["https://*", "http://localhost:*", "http://127.0.0.1:300*", "http://127.0.0.1:864*", "http://127.0.0.1:865*"] });
  await send("Page.addScriptToEvaluateOnNewDocument", { source: `const fixtureQuery = new URLSearchParams(location.search); if (fixtureQuery.has('palette')) localStorage.setItem('omnia.theme', JSON.stringify({palette:fixtureQuery.get('palette'), mode:fixtureQuery.get('mode')}));` });
  await navigate(true); await capture("shot-powerpoint-before.png");
  await navigate(false); await capture("shot-powerpoint-dark-teal.png");
  await evaluate("document.getElementById('settings-button').click()");
  assert(await evaluate("document.getElementById('settings-dialog').open"));
  await capture("shot-powerpoint-settings-dark.png");
  await send("Input.dispatchKeyEvent", { type: "keyDown", key: "Escape", code: "Escape", windowsVirtualKeyCode: 27 });
  await send("Input.dispatchKeyEvent", { type: "keyUp", key: "Escape", code: "Escape", windowsVirtualKeyCode: 27 });
  assert.equal(await evaluate("document.getElementById('settings-dialog').open"), false);
  evidence.interactionChecks.push("Real settings button opens native dialog; Escape dismisses it");
  await navigate(false, "teal", "light", "dark"); await capture("shot-powerpoint-light-teal.png");
  await navigate(false, "gold"); await capture("shot-powerpoint-dark-gold.png");
  for (const width of [320, 380, 450]) {
    for (const palette of ["teal", "blue", "gold", "red"]) {
      for (const mode of ["dark", "light"]) {
        await navigate(false, palette, mode, "dark", width);
        const measured = await metrics();
        assert.equal(measured.width, width); assert.equal(measured.palette, palette); assert.equal(measured.scheme, mode);
        assert(measured.bodyScroll <= width && measured.rootScroll <= width, "Horizontal overflow");
        assert(measured.composer.bottom <= 760 && measured.send.right <= width && measured.messages.height > 0, "Composer clipped");
        evidence.layoutChecks.push(measured);
        await evaluate("document.getElementById('settings-button').click()");
        assert(await evaluate("(() => { const d=document.getElementById('settings-dialog'),r=d.getBoundingClientRect(); return d.open && r.x>=0 && r.right<=innerWidth && r.bottom<=innerHeight && d.scrollWidth<=d.clientWidth; })()"), "Dialog clipped");
        await evaluate("document.getElementById('settings-dialog').close()");
      }
    }
  }
  await navigate(false, "teal", "auto", "light");
  assert.equal((await metrics()).scheme, "light");
  await evaluate("Office.context.officeTheme.isDarkTheme=true; OmniaTheme.refresh()");
  assert.equal((await metrics()).scheme, "dark");
  evidence.interactionChecks.push("Auto follows Office light/dark in built pane");
  await evaluate("document.getElementById('settings-button').click(); document.querySelector('[data-set-palette=red]').click(); document.querySelector('[data-set-mode=light]').click()");
  assert.deepEqual(await evaluate("JSON.parse(localStorage.getItem('omnia.theme'))"), { palette: "red", mode: "light" });
  assert.equal((await metrics()).scheme, "light");
  await evaluate("document.getElementById('close-settings-button').click()");
  assert.equal(await evaluate("document.getElementById('settings-dialog').open"), false);
  evidence.interactionChecks.push("Palette/appearance buttons persist and apply; Close button still dismisses");
  await evaluate("document.getElementById('settings-button').click(); document.getElementById('save-settings-button').click()");
  assert.equal(await evaluate("document.getElementById('settings-dialog').open"), false);
  assert.equal(await evaluate("document.getElementById('connection-status').textContent"), "Settings saved");
  evidence.interactionChecks.push("Connection settings Save behavior preserved");
  assert.equal(exceptions.length, 0, "Browser runtime exceptions");
  await writeFile(path.join(output, "browser-evidence.json"), JSON.stringify(evidence, null, 2));
  console.log(JSON.stringify({ verdict: "PASS", browser: evidence.browser, serverPort: port, screenshots: evidence.captures.map(capture => capture.filename), layoutChecks: evidence.layoutChecks.length, interactionChecks: evidence.interactionChecks, runtimeExceptions: exceptions.length }, null, 2));
} finally {
  if (ws && ws.readyState === WebSocket.OPEN) { try { await send("Browser.close"); } catch { edge.kill(); } ws.close(); }
  else edge.kill();
  server.closeAllConnections();
  await new Promise(resolve => server.close(resolve));
}
