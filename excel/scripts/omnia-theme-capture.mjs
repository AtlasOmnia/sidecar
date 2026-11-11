// Offline-only screenshot harness for the production bundle; never connects to Office/gateway.
import { createServer } from 'node:http';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';

const root = process.cwd();
const out = path.resolve(process.argv[3] || '.tmp/omnia-shots');
const before = process.argv[2] === '--before';
fs.mkdirSync(out, { recursive: true });
fs.mkdirSync('.tmp', { recursive: true });
const server = createServer((req, res) => {
  const file = path.resolve('dist', '.' + new URL(req.url, 'http://localhost').pathname);
  if (!file.startsWith(path.resolve('dist') + path.sep) || !fs.existsSync(file)) {
    res.writeHead(404).end(); return;
  }
  const ext = path.extname(file);
  res.setHeader('Content-Type', ({ '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png' })[ext] || 'application/octet-stream');
  res.end(fs.readFileSync(file));
});
// Request an available high port, never touching 3000-3002.
for (let port = 47140; ; port++) {
  try { await new Promise((resolve, reject) => { server.once('error', reject); server.listen(port, '127.0.0.1', resolve); }); break; }
  catch (e) { if (e.code !== 'EADDRINUSE') throw e; }
}
const origin = `http://127.0.0.1:${server.address().port}`;
const edge = spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', [
  '--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
  '--disable-background-networking', '--disable-extensions', '--remote-debugging-port=0',
  `--user-data-dir=${path.join(root, '.tmp', 'edge-omnia-' + Date.now())}`, 'about:blank'
], { stdio: ['ignore', 'ignore', 'pipe'] });
let socket;
const evidence = { origin, before, scenarios: [], externalRequestsStubbed: [], errors: [] };
try {
  const endpoint = await new Promise((resolve, reject) => {
    let text = '';
    const timer = setTimeout(() => reject(new Error('Edge debugger startup timeout')), 25000);
    edge.stderr.on('data', chunk => { text += chunk; const m = text.match(/DevTools listening on (ws:\/\/\S+)/); if (m) { clearTimeout(timer); resolve(m[1]); } });
    edge.once('error', reject);
  });
  socket = new WebSocket(endpoint);
  await once(socket, 'open');
  let id = 0; const pending = new Map(); const listeners = [];
  function call(method, params = {}, sessionId) {
    return new Promise((resolve, reject) => {
      const key = ++id; const timer = setTimeout(() => { pending.delete(key); reject(new Error('CDP timeout: ' + method)); }, 20000);
      pending.set(key, { resolve, reject, timer });
      socket.send(JSON.stringify({ id: key, method, params, ...(sessionId ? { sessionId } : {}) }));
    });
  }
  socket.addEventListener('message', event => {
    const msg = JSON.parse(event.data);
    if (msg.id) { const p = pending.get(msg.id); if (p) { clearTimeout(p.timer); pending.delete(msg.id); msg.error ? p.reject(new Error(JSON.stringify(msg.error))) : p.resolve(msg.result); } }
    else listeners.forEach(fn => fn(msg));
  });
  const { targetId } = await call('Target.createTarget', { url: 'about:blank' });
  const { sessionId } = await call('Target.attachToTarget', { targetId, flatten: true });
  const cdp = (method, params) => call(method, params, sessionId);
  const evaluate = async expression => {
    const result = await cdp('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
    if (result.exceptionDetails) throw new Error(JSON.stringify(result.exceptionDetails));
    return result.result.value;
  };
  listeners.push(msg => {
    if (msg.method === 'Fetch.requestPaused') {
      const { requestId, request } = msg.params;
      const local = request.url.startsWith(origin + '/');
      if (local) void cdp('Fetch.continueRequest', { requestId });
      else {
        evidence.externalRequestsStubbed.push(request.url);
        const isOffice = request.url.includes('appsforoffice.microsoft.com');
        void cdp('Fetch.fulfillRequest', { requestId, responseCode: isOffice ? 200 : 503,
          responseHeaders: [{ name: 'Content-Type', value: isOffice ? 'text/javascript' : 'application/json' }],
          body: Buffer.from(isOffice ? '/* Office is stubbed by the capture harness. */' : '{"error":"offline screenshot harness"}').toString('base64') });
      }
    }
    if (msg.method === 'Runtime.exceptionThrown') evidence.errors.push(msg.params.exceptionDetails);
  });
  await cdp('Page.enable'); await cdp('Runtime.enable');
  await cdp('Fetch.enable', { patterns: [{ urlPattern: '*', requestStage: 'Request' }] });
  const officeStub = `
    window.Office = { context: { officeTheme: { isDarkTheme: true }, document: { settings: {
      get: () => '7d58038f-619c-4bc6-8cf6-e1b7a28b2165', set: () => {}, saveAsync: cb => cb({status:'succeeded'})
    } } }, AsyncResultStatus: { Succeeded: 'succeeded' }, onReady: cb => {
      if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => cb({host:'Excel'}), {once:true});
      else queueMicrotask(() => cb({host:'Excel'}));
    } };
    if (!localStorage.getItem('omnia.theme')) localStorage.setItem('omnia.theme', JSON.stringify({ palette: 'teal', mode: 'dark' }));
  `;
  await cdp('Page.addScriptToEvaluateOnNewDocument', { source: officeStub });
  await cdp('Emulation.setDeviceMetricsOverride', { width: 380, height: 760, deviceScaleFactor: 1, mobile: false });
  await cdp('Page.navigate', { url: origin + '/taskpane.html' });
  for (let attempt = 0; attempt < 100; attempt++) {
    if (await evaluate(`document.getElementById('connection-status')?.textContent === 'Ready'`)) break;
    if (attempt === 99) throw new Error('Built taskpane failed to boot');
    await new Promise(r => setTimeout(r, 100));
  }
  // Explicitly synthetic display fixture, not a claim of workbook/model execution.
  await evaluate(`(() => {
    document.getElementById('messages').innerHTML = '<article class="message user"><div class="bubble">Summarize this month’s budget.</div></article><article class="message assistant"><div class="bubble"><p>Here’s a <strong>clear view</strong> of your budget.</p><ul><li>Revenue is ahead of plan.</li><li>Operating costs remain steady.</li></ul><p>I can help you explore the details or build a chart next.</p></div><details class="activity" open><summary>Activity</summary><div class="activity-list"><div class="activity-item"><span>✓</span><span>Read Sheet1!A1:D12</span></div></div></details></article>';
    document.getElementById('bridge-status').textContent = 'Office bridge connected';
    document.getElementById('bridge-status').className = 'bridge-status bridge-connected';
    document.getElementById('bridge-session-id').textContent = '7d58038f-619c-4bc6-8cf6-e1b7a28b2165';
    document.getElementById('copy-bridge-session').disabled = false;
  })()`);
  const settle = () => evaluate('new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))');
  const shot = async (name, settings = false) => {
    await settle();
    const metrics = await evaluate(`(() => {
      const r = document.getElementById('send-button').getBoundingClientRect();
      return { width:innerWidth, height:innerHeight, scrollWidth:document.documentElement.scrollWidth, palette:document.documentElement.dataset.palette, scheme:document.documentElement.dataset.scheme, dialog:document.getElementById('settings-dialog').open, composerVisible:r.top>=0 && r.bottom<=innerHeight, background:getComputedStyle(document.body).backgroundColor };
    })()`);
    assert.ok(metrics.scrollWidth <= metrics.width, 'Horizontal overflow');
    assert.ok(settings || metrics.composerVisible, 'Composer clipped');
    const { data } = await cdp('Page.captureScreenshot', { format: 'png', captureBeyondViewport: false });
    fs.writeFileSync(path.join(out, name), Buffer.from(data, 'base64'));
    evidence.scenarios.push({ name, ...metrics });
  };
  if (before) await shot('shot-excel-before.png');
  else {
    await shot('shot-excel-dark-teal.png');
    await evaluate(`document.getElementById('settings-button').click(); document.querySelector('[name="omnia-mode"][value="light"]').click()`);
    await evaluate(`document.getElementById('close-settings-button').click()`);
    await shot('shot-excel-light-teal.png');
    await evaluate(`document.getElementById('settings-button').click(); document.querySelector('[name="omnia-mode"][value="dark"]').click()`);
    await shot('shot-excel-settings-dark.png', true);
    await evaluate(`document.querySelector('[name="omnia-palette"][value="gold"]').click(); document.getElementById('close-settings-button').click()`);
    await shot('shot-excel-dark-gold.png');
    // Exercise live Office Auto + fallback and OS changes, and responsive/keyboard controls.
    await evaluate(`document.getElementById('settings-button').click(); document.querySelector('[name="omnia-mode"][value="auto"]').click(); document.getElementById('close-settings-button').click(); Office.context.officeTheme = {isDarkTheme:false}`);
    await new Promise(r => setTimeout(r, 1100));
    assert.equal(await evaluate('document.documentElement.dataset.scheme'), 'light');
    await evaluate('Office.context.officeTheme = {isDarkTheme:true}');
    await new Promise(r => setTimeout(r, 1100));
    assert.equal(await evaluate('document.documentElement.dataset.scheme'), 'dark');
    evidence.officeAutoLive = true;
    await evaluate('Office.context.officeTheme = undefined');
    for (const value of ['light', 'dark']) {
      await cdp('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-color-scheme', value }] });
      await new Promise(r => setTimeout(r, 1100));
      assert.equal(await evaluate('document.documentElement.dataset.scheme'), value);
    }
    evidence.osFallbackLive = true;
    await evaluate(`Office.context.officeTheme = {bodyBackgroundColor:'#ffffff'}`);
    await new Promise(r => setTimeout(r, 1100));
    assert.equal(await evaluate('document.documentElement.dataset.scheme'), 'light');
    await evaluate('Office.context.officeTheme = {isDarkTheme:true}');
    await new Promise(r => setTimeout(r, 1100));
    for (const width of [320, 380, 450]) {
      await cdp('Emulation.setDeviceMetricsOverride', { width, height: 760, deviceScaleFactor: 1, mobile: false });
      await settle();
      assert.equal(await evaluate('document.documentElement.scrollWidth <= innerWidth'), true);
      await evaluate(`document.getElementById('settings-button').click()`);
      assert.equal(await evaluate(`document.getElementById('save-settings-button').getBoundingClientRect().bottom <= innerHeight`), true);
      await evaluate(`document.getElementById('close-settings-button').click()`);
    }
    evidence.responsiveWidths = [320, 380, 450];
    await evaluate(`document.getElementById('settings-button').click(); document.querySelector('[name="omnia-palette"][value="gold"]').focus()`);
    await cdp('Input.dispatchKeyEvent', { type: 'keyDown', key: 'ArrowRight', code: 'ArrowRight', windowsVirtualKeyCode: 39 });
    await cdp('Input.dispatchKeyEvent', { type: 'keyUp', key: 'ArrowRight', code: 'ArrowRight', windowsVirtualKeyCode: 39 });
    assert.equal(await evaluate('document.documentElement.dataset.palette'), 'red');
    evidence.nativeRadioKeyboard = true;
    await evaluate(`document.querySelector('[name="omnia-palette"][value="gold"]').click()`);
    await cdp('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
    await cdp('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
    assert.equal(await evaluate(`document.getElementById('settings-dialog').open`), false);
    evidence.escapeDismissesSettings = true;
    const stored = await evaluate(`JSON.parse(localStorage.getItem('omnia.theme'))`);
    assert.deepEqual(stored, { palette: 'gold', mode: 'auto' });
    evidence.persisted = stored;
    evidence.paletteContrast = [];
    for (const palette of ['teal', 'blue', 'gold', 'red']) {
      for (const mode of ['dark', 'light']) {
        await evaluate(`document.getElementById('settings-button').click(); document.querySelector('[name="omnia-palette"][value="${palette}"]').click(); document.querySelector('[name="omnia-mode"][value="${mode}"]').click(); document.getElementById('close-settings-button').click()`);
        const colors = await evaluate(`(() => {
          const button = getComputedStyle(document.getElementById('send-button'));
          return { fg:button.color, bg:button.backgroundColor };
        })()`);
        const luminance = rgb => rgb.match(/[\d.]+/g).slice(0,3).map(Number).map(n => { n /= 255; return n <= .04045 ? n / 12.92 : ((n + .055) / 1.055) ** 2.4; }).reduce((sum, n, i) => sum + n * [.2126,.7152,.0722][i], 0);
        const fg = luminance(colors.fg), bg = luminance(colors.bg);
        const ratio = (Math.max(fg,bg) + .05) / (Math.min(fg,bg) + .05);
        assert.ok(ratio >= 4.5, `${palette}/${mode} Send contrast: ${ratio}`);
        evidence.paletteContrast.push({ palette, mode, ...colors, ratio });
      }
    }
    await evaluate(`document.getElementById('settings-button').click(); document.querySelector('[name="omnia-palette"][value="gold"]').click(); document.querySelector('[name="omnia-mode"][value="auto"]').click(); document.getElementById('close-settings-button').click()`);
    await cdp('Page.reload');
    for (let attempt = 0; attempt < 100; attempt++) {
      if (await evaluate(`document.getElementById('connection-status')?.textContent === 'Ready'`)) break;
      if (attempt === 99) throw new Error('Reload failed');
      await new Promise(r => setTimeout(r, 100));
    }
    assert.equal(await evaluate('document.documentElement.dataset.palette'), 'gold');
    assert.equal(await evaluate('document.documentElement.dataset.scheme'), 'dark');
    assert.deepEqual(await evaluate(`JSON.parse(localStorage.getItem('omnia.theme'))`), stored);
    evidence.persistenceAfterReload = true;
  }
  assert.equal(evidence.errors.length, 0, 'Browser runtime exceptions');
  fs.writeFileSync(path.join(out, before ? 'capture-before.json' : 'capture-after.json'), JSON.stringify(evidence, null, 2));
  console.log(JSON.stringify(evidence, null, 2));
  await call('Browser.close');
} finally {
  if (socket) socket.close();
  edge.kill();
  server.close();
}
