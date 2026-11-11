/* Static built-pane acceptance. No Office/relay/live services are contacted.
   Setup: npm install --prefix .tmp/browser --no-package-lock --ignore-scripts playwright-core
   Run after npm run build, with TMP/TEMP/TMPDIR beneath this worktree:
     node scripts/capture-profile-pane.cjs <absolute-output-directory>
*/
const { chromium } = require('../.tmp/browser/node_modules/playwright-core');
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '../dist');
const output = path.resolve(process.argv[2] || '.tmp/screenshots');
const mime = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png' };
const server = http.createServer((req, res) => {
  const target = path.resolve(root, '.' + new URL(req.url, 'http://fixture').pathname);
  if (!target.startsWith(root + path.sep) || !fs.existsSync(target)) { res.writeHead(404).end(); return; }
  res.setHeader('Content-Type', mime[path.extname(target)] || 'application/octet-stream');
  res.end(fs.readFileSync(target));
});
(async () => {
  fs.mkdirSync(output, { recursive: true });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const origin = `http://127.0.0.1:${server.address().port}`;
  let browser;
  try {
    browser = await chromium.launch({ executablePath: 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', headless: true, args: ['--disable-gpu'] });
    const context = await browser.newContext({ viewport: { width: 240, height: 720 }, deviceScaleFactor: 1 });
    await context.route('**/*', route => {
      if (route.request().url().startsWith(origin + '/')) return route.continue();
      if (route.request().url().includes('/office.js')) return route.fulfill({ body: '', contentType: 'text/javascript' });
      return route.abort();
    });
    await context.addInitScript(() => {
      // Faithful protocol fakes, intentionally not a simulated real Office result.
      const state = window.fixture = { generation: 7, label: 'Mac · demo', legacy: false, status: 200, session: 0, payloads: [] };
      const settings = new Map();
      window.Office = {
        onReady: cb => cb(), AsyncResultStatus: { Succeeded: 'succeeded' },
        context: { requirements: { isSetSupported: () => false }, document: { settings: {
          get: k => settings.get(k), set: (k, v) => settings.set(k, v), saveAsync: cb => cb({ status: 'succeeded' }),
        } } },
      };
      window.Word = { run: async cb => cb({
        document: { body: { text: 'Static fixture document.', load() {}, paragraphs: { items: [{}], load() {} } },
          getSelection: () => ({ text: '', load() {} }), }, sync: async () => {},
      }) };
      window.fetch = async (url, options = {}) => {
        const pathname = new URL(String(url), location.href).pathname;
        const json = (value, status = 200, headers = {}) => new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json', ...headers } });
        const headers = () => state.legacy ? {} : { 'X-Omnia-Profile': 'demo', 'X-Omnia-Profile-Label': encodeURIComponent(state.label), 'X-Omnia-Profile-Generation': String(state.generation) };
        if (pathname.endsWith('/omnia/profile')) return state.legacy ? new Response(null, { status: 404 }) : json({ version: 1, connectionId: 'fixture', connectionLabel: 'Mac', profile: 'demo', profileLabel: state.label, generation: state.generation }, 200, headers());
        if (pathname.endsWith('/chat/completions')) {
          state.payloads.push({ body: JSON.parse(options.body), session: new Headers(options.headers).get('X-Hermes-Session-Id') });
          return json({ choices: [{ message: { content: 'Fixture reply: ready to help with this document.' } }] }, state.status, headers());
        }
        if (pathname === '/word/session') return options.method === 'DELETE' ? json({}) : json({ protocolVersion: '1.0', sessionId: `660e8400-e29b-41d4-a716-${String(++state.session).padStart(12, '0')}`, sessionKey: 'disposable-fixture-session-key-only', expiresInSeconds: 60 }, 201);
        if (pathname === '/word/pending') return new Response(null, { status: 204 });
        if (pathname === '/word/heartbeat') return json({});
        throw new Error('Unexpected fixture request: ' + pathname);
      };
    });
    const page = await context.newPage();
    const errors = []; page.on('pageerror', e => errors.push(e.message));
    const screenshots = [];
    async function capture(name) {
      assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true, 'narrow pane overflow');
      await page.screenshot({ path: path.join(output, name + '.png'), fullPage: false, animations: 'disabled' });
      screenshots.push(name + '.png');
    }
    async function send(text) {
      const expected = await page.evaluate(() => window.fixture.payloads.length + 1);
      await page.fill('#prompt-input', text); await page.click('#send-button');
      await page.waitForFunction(count => !document.querySelector('#send-button').disabled && window.fixture.payloads.length === count, expected);
    }
    await page.goto(origin + '/taskpane.html');
    await page.waitForSelector('.omnia-profile-chip:not([hidden])');
    assert.equal(await page.getAttribute('.omnia-profile-chip', 'aria-label'), 'Talking to Mac · demo');
    await capture('01-chip-narrow');
    await send('old profile private question');
    await page.fill('#prompt-input', 'Keep my unsent draft');
    await page.evaluate(() => { window.fixture.generation = 8; window.fixture.label = 'Mac · other'; dispatchEvent(new Event('focus')); });
    await page.waitForSelector('.omnia-profile-note');
    assert.equal(await page.inputValue('#prompt-input'), 'Keep my unsent draft');
    assert.equal(await page.textContent('.omnia-profile-note'), 'Now talking to Mac · other');
    await capture('02-now-talking-narrow');
    await send('fresh new-profile question');
    const payloads = await page.evaluate(() => window.fixture.payloads);
    assert.equal(payloads.length, 2);
    assert.equal(payloads[1].body.messages.length, 2);
    assert.ok(!JSON.stringify(payloads[1].body).includes('old profile private question'));
    assert.ok(!JSON.stringify(payloads[1].body).includes('Now talking to'));
    assert.notEqual(payloads[1].session, payloads[0].session);
    await page.evaluate(() => { window.fixture.status = 409; });
    await send('profile error fixture');
    await page.waitForFunction(() => document.querySelector('#messages').textContent.includes("Omnia Link can't reach Mac · other right now. Check Omnia Link."));
    assert.ok(!(await page.textContent('#messages')).includes('409'));
    await capture('03-plain-error-narrow');
    await page.reload();
    await page.waitForSelector('.omnia-profile-chip:not([hidden])');
    await page.evaluate(() => { window.fixture.legacy = true; dispatchEvent(new Event('focus')); });
    await page.waitForSelector('.omnia-profile-chip[hidden]', { state: 'attached' });
    await send('legacy still works');
    assert.equal(await page.isVisible('.omnia-profile-chip'), false);
    await capture('04-old-bridge-narrow');
    assert.deepEqual(errors, []);
    fs.writeFileSync(path.join(output, 'acceptance.json'), JSON.stringify({ viewport: { width: 240, height: 720 }, builtPane: true, fakeOfficeAndRelay: true, liveServicesContacted: false, screenshots, assertionsPassed: true, browser: await browser.version(), errors }, null, 2));
    console.log(JSON.stringify({ screenshots, assertionsPassed: true, output }, null, 2));
  } finally {
    if (browser) await browser.close();
    await new Promise(resolve => server.close(resolve));
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
