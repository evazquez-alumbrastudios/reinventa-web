// node --test tests/   (solo datos ficticios; no usa red)
const test = require('node:test'), assert = require('node:assert'), fs = require('fs'), vm = require('vm'), path = require('path');
const root = path.join(__dirname, '..');
const FICT = 'clave-ficticia-de-prueba', PINF = '123456';
let calls = [];
global.fetch = async (url, o) => { calls.push({ url: String(url), body: o && o.body }); return { ok: true, json: async () => ({ status: 'ok' }) }; };
async function call(name, req) {
  const h = require(path.join(root, 'api', name));
  const res = { code: 200, setHeader() {}, status(c) { this.code = c; return this; }, json(b) { this.body = b; return this; }, end() { return this; } };
  await h({ method: 'GET', query: {}, headers: {}, body: {}, ...req }, res); return res;
}
const env = (a, p) => { calls = []; a ? (process.env.ADMIN_KEY = a) : delete process.env.ADMIN_KEY; p ? (process.env.STAFF_PIN = p) : delete process.env.STAFF_PIN; };
const admin = (k, q = {}) => call('admin.js', { headers: k ? { 'x-admin-key': k } : {}, query: q });
const action = (key) => call('admin-action.js', { method: 'POST', body: { key, sub: 'enviar_agradecimiento_masivo' } });
const entrada = (pin) => call('entrada.js', { method: 'POST', body: { id: 'FICT-001', pin } });

test('sin variables todo queda cerrado y no llama al Apps Script', async () => {
  env();
  for (const r of [await admin('x'), await admin(), await action('x'), await action(), await entrada('0000'), await entrada('')]) assert.equal(r.code, 503);
  assert.equal(calls.length, 0);
});
test('clave o PIN incorrectos se rechazan sin llamar al Apps Script', async () => {
  env(FICT, PINF);
  for (const r of [await admin('otra'), await admin(), await admin(undefined, { key: FICT }), await action('otra'), await action()]) assert.equal(r.code, 401);
  assert.deepEqual((await entrada('999999')).body, { error: 'pin_incorrecto' });
  assert.equal(calls.length, 0);
});
test('con la clave correcta no se reenvía la clave al Apps Script', async () => {
  env(FICT, PINF);
  assert.equal((await admin(FICT, { ping: '1' })).code, 200); assert.equal(calls.length, 0);
  await action(FICT); assert.equal(calls.length, 1); assert.ok(!calls[0].body.includes(FICT));
});
test('no quedan credenciales por defecto en el código', () => {
  const files = [...fs.readdirSync(path.join(root, 'api')).filter((f) => f.endsWith('.js')).map((f) => 'api/' + f), 'google-sheets-integration.gs', 'admin.html', 'entrada.html'];
  for (const f of files) {
    const s = fs.readFileSync(path.join(root, f), 'utf8');
    assert.ok(!/DEFAULT_(KEY|PIN)|process\.env\.(ADMIN_KEY|STAFF_PIN)\s*\|\||var\s+STAFF_PIN|\/api\/admin\?key=/.test(s), f);
  }
});
test('el Apps Script del repo sigue bloqueando administración aun con clave correcta', () => {
  let touched = 0;
  const ctx = { ContentService: { MimeType: {}, createTextOutput: (t) => ({ text: t, setMimeType() { return this; } }) },
    SpreadsheetApp: { getActiveSpreadsheet() { touched++; throw new Error('no'); } }, MailApp: { sendEmail() { touched++; } } };
  vm.createContext(ctx); vm.runInContext(fs.readFileSync(path.join(root, 'google-sheets-integration.gs'), 'utf8'), ctx);
  for (const a of ['admin', 'hub', 'entrada']) assert.equal(ctx.doGet({ parameter: { action: a, id: 'FICT-001' } }).text, '{"status":"ok"}');
  const subs = [...fs.readFileSync(path.join(root, 'google-sheets-integration.gs'), 'utf8').matchAll(/sub === '([a-z_]+)'/g)].map((m) => m[1]);
  assert.ok(subs.length >= 40);
  for (const sub of subs) assert.match(ctx.doPost({ postData: { contents: JSON.stringify({ action: 'admin_action', sub }) } }).text, /deshabilitado/);
  assert.equal(touched, 0);
});
