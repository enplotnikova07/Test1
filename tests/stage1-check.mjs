import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { makeApplication } from '../server.mjs';

const directory = await mkdtemp(join(tmpdir(), 'zapasometr-stage1-'));
const { server, store } = makeApplication({ databasePath: join(directory, 'test.db'), secureCookies: false });
makeApplication({ databasePath: join(directory, 'test.db'), secureCookies: false });
await new Promise((resolve) => server.listen(0, resolve));
const origin = `http://127.0.0.1:${server.address().port}`;
const close = () => new Promise((resolve) => server.close(resolve));

async function call(path, { method = 'GET', body, cookie, csrf } = {}) {
  const response = await fetch(`${origin}${path}`, { method, headers: { ...(body ? { 'content-type': 'application/json' } : {}), ...(cookie ? { cookie } : {}), ...(csrf ? { 'x-csrf-token': csrf } : {}) }, body: body ? JSON.stringify(body) : undefined });
  const text = await response.text();
  let data = {}; try { data = JSON.parse(text); } catch { data = { text }; }
  return { response, data, cookie: response.headers.get('set-cookie')?.split(';')[0] };
}

try {
  const health = await call('/api/health');
  assert.equal(health.response.status, 200); assert.equal(health.data.ok, true);
  const page = await call('/');
  assert.equal(page.response.status, 200); assert.match(page.data.text, /Управление товарными остатками/);

  const admin = await store.createUser({ email: 'admin@example.test', name: 'Администратор', password: 'ChangeMe!2026', role: 'ADMIN' });
  const login = await call('/api/auth/login', { method: 'POST', body: { email: admin.email, password: 'ChangeMe!2026' } });
  assert.equal(login.response.status, 200); assert.ok(login.cookie); assert.ok(login.data.csrfToken);
  assert.match(login.response.headers.get('set-cookie'), /HttpOnly/); assert.match(login.response.headers.get('set-cookie'), /SameSite=Lax/);

  const csrfRejected = await call('/api/users', { method: 'POST', cookie: login.cookie, body: { email: 'ignored@example.test', name: 'Без токена', password: 'SomePass!2026', role: 'BUYER' } });
  assert.equal(csrfRejected.response.status, 403);
  const created = await call('/api/users', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { email: 'buyer@example.test', name: 'Закупщик', password: 'BuyerPass!2026', role: 'BUYER' } });
  assert.equal(created.response.status, 201); assert.equal(created.data.user.role, 'BUYER');
  const buyer = created.data.user;
  const buyerLogin = await call('/api/auth/login', { method: 'POST', body: { email: buyer.email, password: 'BuyerPass!2026' } });
  const forbidden = await call('/api/users', { cookie: buyerLogin.cookie }); assert.equal(forbidden.response.status, 403);
  for (const [role, email, name] of [['MANAGER', 'manager@example.test', 'Руководитель'], ['FINANCE', 'finance@example.test', 'Финансовый директор']]) {
    const user = await store.createUser({ email, name, password: 'RolePass!2026', role }, admin.id);
    const roleLogin = await call('/api/auth/login', { method: 'POST', body: { email: user.email, password: 'RolePass!2026' } });
    assert.equal(roleLogin.response.status, 200); assert.equal(roleLogin.data.user.role, role);
    const roleForbidden = await call('/api/users', { cookie: roleLogin.cookie }); assert.equal(roleForbidden.response.status, 403);
  }
  const adminCannotSeeBusiness = await call('/api/products', { cookie: login.cookie }); assert.equal(adminCannotSeeBusiness.response.status, 404);

  const preventLastAdmin = await call(`/api/users/${admin.id}/block`, { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken }); assert.equal(preventLastAdmin.response.status, 400);
  const blocked = await call(`/api/users/${buyer.id}/block`, { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken }); assert.equal(blocked.response.status, 200); assert.equal(blocked.data.user.active, false);
  const blockedSession = await call('/api/auth/me', { cookie: buyerLogin.cookie }); assert.equal(blockedSession.response.status, 401);
  const unblocked = await call(`/api/users/${buyer.id}/unblock`, { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken }); assert.equal(unblocked.response.status, 200); assert.equal(unblocked.data.user.active, true);

  const reset = await call(`/api/users/${buyer.id}/password-reset`, { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken }); assert.equal(reset.response.status, 201); assert.match(reset.data.resetUrl, /reset=/);
  const token = new URL(reset.data.resetUrl).searchParams.get('reset');
  const resetDone = await call('/api/auth/reset-password', { method: 'POST', body: { token, password: 'UpdatedPass!2026' } }); assert.equal(resetDone.response.status, 200);
  const newPassword = await call('/api/auth/login', { method: 'POST', body: { email: buyer.email, password: 'UpdatedPass!2026' } }); assert.equal(newPassword.response.status, 200);
  const logout = await call('/api/auth/logout', { method: 'POST', cookie: newPassword.cookie, csrf: newPassword.data.csrfToken }); assert.equal(logout.response.status, 200);
  const loggedOut = await call('/api/auth/me', { cookie: newPassword.cookie }); assert.equal(loggedOut.response.status, 401);

  const audit = await call('/api/audit', { cookie: login.cookie }); assert.equal(audit.response.status, 200);
  for (const action of ['USER_CREATED', 'USER_BLOCKED', 'USER_UNBLOCKED', 'PASSWORD_RESET_REQUESTED']) assert.ok(audit.data.events.some((event) => event.action === action), `Нет события ${action}`);

  const expiringLogin = await call('/api/auth/login', { method: 'POST', body: { email: admin.email, password: 'ChangeMe!2026' } });
  const expiringSessionId = expiringLogin.cookie.split('=')[1];
  store.db.prepare('UPDATE sessions SET expires_at = ? WHERE id = ?').run(new Date(Date.now() - 1000).toISOString(), expiringSessionId);
  const expired = await call('/api/auth/me', { cookie: expiringLogin.cookie }); assert.equal(expired.response.status, 401);

  const rateDirectory = await mkdtemp(join(tmpdir(), 'zapasometr-rate-'));
  const rateApp = makeApplication({ databasePath: join(rateDirectory, 'test.db'), secureCookies: false }); await rateApp.store.createUser({ email: 'rate@example.test', name: 'Проверка', password: 'ChangeMe!2026', role: 'ADMIN' }); await new Promise((resolve) => rateApp.server.listen(0, resolve));
  const rateOrigin = `http://127.0.0.1:${rateApp.server.address().port}`;
  for (let attempt = 0; attempt < 5; attempt += 1) { const response = await fetch(`${rateOrigin}/api/auth/login`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ email: 'rate@example.test', password: 'wrong' }) }); assert.equal(response.status, 401); }
  const locked = await fetch(`${rateOrigin}/api/auth/login`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ email: 'rate@example.test', password: 'ChangeMe!2026' }) }); assert.equal(locked.status, 429);
  await new Promise((resolve) => rateApp.server.close(resolve));

  console.log('Этап 1: 33 автоматические проверки пройдены.');
} finally {
  await close();
}
