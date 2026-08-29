import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { makeApplication } from '../server.mjs';

async function appFixture() {
  const directory = await mkdtemp(join(tmpdir(), 'zapasometr-test-'));
  const { server, store } = makeApplication({ databasePath: join(directory, 'test.db'), secureCookies: false });
  await new Promise((resolve) => server.listen(0, resolve));
  const origin = `http://127.0.0.1:${server.address().port}`;
  return { origin, store, close: () => new Promise((resolve) => server.close(resolve)) };
}
async function call(origin, path, { method = 'GET', body, cookie, csrf } = {}) {
  const response = await fetch(`${origin}${path}`, { method, headers: { ...(body ? { 'content-type': 'application/json' } : {}), ...(cookie ? { cookie } : {}), ...(csrf ? { 'x-csrf-token': csrf } : {}) }, body: body ? JSON.stringify(body) : undefined });
  return { response, data: await response.json(), cookie: response.headers.get('set-cookie')?.split(';')[0] };
}

test('администратор создаёт пользователя, а бизнес-пользователь не получает административный доступ', async (t) => {
  const fixture = await appFixture(); t.after(fixture.close);
  await fixture.store.createUser({ email: 'admin@example.test', name: 'Администратор', password: 'ChangeMe!2026', role: 'ADMIN' });
  const login = await call(fixture.origin, '/api/auth/login', { method: 'POST', body: { email: 'admin@example.test', password: 'ChangeMe!2026' } });
  assert.equal(login.response.status, 200); assert.ok(login.cookie); assert.ok(login.data.csrfToken);
  const created = await call(fixture.origin, '/api/users', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { email: 'buyer@example.test', name: 'Закупщик', password: 'BuyerPass!2026', role: 'BUYER' } });
  assert.equal(created.response.status, 201); assert.equal(created.data.user.role, 'BUYER');
  const buyerLogin = await call(fixture.origin, '/api/auth/login', { method: 'POST', body: { email: 'buyer@example.test', password: 'BuyerPass!2026' } });
  const forbidden = await call(fixture.origin, '/api/users', { cookie: buyerLogin.cookie });
  assert.equal(forbidden.response.status, 403);
  const audit = await call(fixture.origin, '/api/audit', { cookie: login.cookie });
  assert.equal(audit.response.status, 200); assert.ok(audit.data.events.some((event) => event.action === 'USER_CREATED'));
});

test('нельзя блокировать последнего администратора, блокировка завершает сессии', async (t) => {
  const fixture = await appFixture(); t.after(fixture.close);
  const admin = await fixture.store.createUser({ email: 'admin@example.test', name: 'Администратор', password: 'ChangeMe!2026', role: 'ADMIN' });
  const buyer = await fixture.store.createUser({ email: 'buyer@example.test', name: 'Закупщик', password: 'BuyerPass!2026', role: 'BUYER' }, admin.id);
  const adminLogin = await call(fixture.origin, '/api/auth/login', { method: 'POST', body: { email: admin.email, password: 'ChangeMe!2026' } });
  const prevent = await call(fixture.origin, `/api/users/${admin.id}/block`, { method: 'POST', cookie: adminLogin.cookie, csrf: adminLogin.data.csrfToken });
  assert.equal(prevent.response.status, 400);
  const buyerLogin = await call(fixture.origin, '/api/auth/login', { method: 'POST', body: { email: buyer.email, password: 'BuyerPass!2026' } });
  const blocked = await call(fixture.origin, `/api/users/${buyer.id}/block`, { method: 'POST', cookie: adminLogin.cookie, csrf: adminLogin.data.csrfToken });
  assert.equal(blocked.response.status, 200);
  const current = await call(fixture.origin, '/api/auth/me', { cookie: buyerLogin.cookie });
  assert.equal(current.response.status, 401);
});

test('неверные входы блокируются после пяти попыток, а CSRF обязателен для изменений', async (t) => {
  const fixture = await appFixture(); t.after(fixture.close);
  await fixture.store.createUser({ email: 'admin@example.test', name: 'Администратор', password: 'ChangeMe!2026', role: 'ADMIN' });
  const login = await call(fixture.origin, '/api/auth/login', { method: 'POST', body: { email: 'admin@example.test', password: 'ChangeMe!2026' } });
  const csrf = await call(fixture.origin, '/api/users', { method: 'POST', cookie: login.cookie, body: { email: 'missing@example.test', name: 'Без токена', password: 'SomePass!2026', role: 'BUYER' } });
  assert.equal(csrf.response.status, 403);
  for (let attempt = 0; attempt < 5; attempt += 1) { const result = await call(fixture.origin, '/api/auth/login', { method: 'POST', body: { email: 'admin@example.test', password: 'bad-password' } }); assert.equal(result.response.status, 401); }
  const locked = await call(fixture.origin, '/api/auth/login', { method: 'POST', body: { email: 'admin@example.test', password: 'ChangeMe!2026' } });
  assert.equal(locked.response.status, 429);
});
