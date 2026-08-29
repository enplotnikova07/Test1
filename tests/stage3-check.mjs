import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { makeApplication } from '../server.mjs';

const directory = await mkdtemp(join(tmpdir(), 'inventory-analysis-'));
const { server, store } = makeApplication({ databasePath: join(directory, 'test.db'), secureCookies: false });
await new Promise((resolve) => server.listen(0, resolve));
const origin = `http://127.0.0.1:${server.address().port}`;
const close = () => new Promise((resolve) => server.close(resolve));
const source = [
  'Номенклатура;1;2;3;4;5;6;7;8;Тек.остаток, шт;Средняя Цена 1 единицы;Средняя Себестоимость 1 единицы;За сколько нам приходит товар от поставщика, дн',
  'Нет в наличии;10;10;10;10;10;10;10;10;0;100;60;5',
  'Неликвид;0;0;0;0;0;0;0;0;20;50;30;5',
  'Заказать срочно;10;10;10;10;10;10;10;10;5;80;50;5',
  'Заказать скоро;10;10;10;10;10;10;10;10;10;80;50;5',
  'Нерегулярный спрос;0;20;0;20;0;20;0;20;10;70;45;5',
  'Избыточный запас;2;2;2;2;2;2;2;2;300;30;10;5',
  'Снижение активности;1;0;0;0;0;0;0;0;50;30;10;5',
  'Недостаточно данных;1;1;1;1;1;1;1;1;50;30;10;5'
].join('\n');
async function call(path, { method = 'GET', body, cookie, csrf } = {}) { const response = await fetch(`${origin}${path}`, { method, headers: { ...(body ? { 'content-type': 'application/json' } : {}), ...(cookie ? { cookie } : {}), ...(csrf ? { 'x-csrf-token': csrf } : {}) }, body: body ? JSON.stringify(body) : undefined }); const text = await response.text(); return { response, data: text ? JSON.parse(text) : {}, cookie: response.headers.get('set-cookie')?.split(';')[0] }; }

try {
  const buyer = await store.createUser({ email: 'buyer@example.test', name: 'Закупщик', password: 'BuyerPass!2026', role: 'BUYER' });
  const login = await call('/api/auth/login', { method: 'POST', body: { email: buyer.email, password: 'BuyerPass!2026' } });
  const preview = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'PASTE', text: source } } });
  const imported = await call('/api/imports/commit', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { previewId: preview.data.previewId, mapping: preview.data.suggested.mapping, periodColumns: preview.data.suggested.periodColumns, period: { granularity: 'week', lastPeriodEnd: '2026-08-24', order: 'OLDEST_TO_NEWEST' } } });
  assert.equal(imported.response.status, 201); assert.equal(imported.data.analysis.rows, 8);
  const resultResponse = await call('/api/analytics/results', { cookie: login.cookie }); assert.equal(resultResponse.response.status, 200); assert.equal(resultResponse.data.results.length, 8);
  const byName = Object.fromEntries(resultResponse.data.results.map((row) => [row.product_name, row]));
  assert.equal(byName['Нет в наличии'].primary_status, 'NO_STOCK'); assert.equal(byName.Неликвид.primary_status, 'OBSOLETE'); assert.equal(byName['Заказать срочно'].primary_status, 'ORDER_NOW'); assert.equal(byName['Избыточный запас'].primary_status, 'EXCESS');
  assert.equal(byName['Нет в наличии'].xyz, 'X'); assert.equal(byName['Нерегулярный спрос'].xyz, 'Z'); assert.equal(byName.Неликвид.xyz, null);
  assert.equal(byName['Нет в наличии'].abc_quantity, 'A'); assert.ok(byName['Нет в наличии'].reorder_point > 7 && byName['Нет в наличии'].reorder_point < 8); assert.equal(byName['Нет в наличии'].recommended_qty, 18);
  assert.ok(byName['Нерегулярный спрос'].cv > 90); assert.ok(byName['Нерегулярный спрос'].safety_stock > 0);
  const settingsBefore = await call('/api/settings', { cookie: login.cookie }); assert.equal(settingsBefore.data.active.params.safetyZ, 1.65);
  const invalidSettings = await call('/api/settings', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { params: { abcA: 96, abcB: 95 } } }); assert.equal(invalidSettings.response.status, 400);
  const changedSettings = await call('/api/settings', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { params: { abcA: 60, abcB: 85, xyzX: 5, xyzY: 20, safetyZ: 2, slowDays: 20, obsoleteDays: 80, excessDays: 9999, minHistory: 12, excludeOutliers: true } } }); assert.equal(changedSettings.response.status, 201); assert.equal(changedSettings.data.analysis.rows, 8);
  const settingsAfter = await call('/api/settings', { cookie: login.cookie }); assert.equal(settingsAfter.data.active.params.safetyZ, 2); assert.equal(settingsAfter.data.history.length, 2);
  const changedResults = await call('/api/analytics/results', { cookie: login.cookie }); const changedByName = Object.fromEntries(changedResults.data.results.map((row) => [row.product_name, row]));
  assert.equal(changedByName['Заказать скоро'].primary_status, 'ORDER_SOON'); assert.equal(changedByName['Снижение активности'].primary_status, 'SLOW'); assert.equal(changedByName['Недостаточно данных'].primary_status, 'INSUFFICIENT_DATA');
  const recalculated = await call('/api/analytics/recalculate', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken }); assert.equal(recalculated.response.status, 200); assert.equal(recalculated.data.analysis.rows, 8);
  const audit = store.listAudit(); assert.ok(audit.some((event) => event.action === 'SETTINGS_UPDATED')); assert.ok(audit.some((event) => event.action === 'ANALYSIS_RECALCULATED'));
  console.log('Этап 3 расчёты: 31 автоматическая проверка пройдена.');
} finally { await close(); }
