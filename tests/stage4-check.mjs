import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { makeApplication } from '../server.mjs';

const directory = await mkdtemp(join(tmpdir(), 'inventory-screens-'));
const { server, store } = makeApplication({ databasePath: join(directory, 'test.db'), secureCookies: false });
await new Promise((resolve) => server.listen(0, resolve));
const origin = `http://127.0.0.1:${server.address().port}`;
const source = ['Номенклатура;Характеристика;Код 1С;Категория;Поставщик;1;2;3;4;5;6;7;8;Тек.остаток, шт;Средняя Цена 1 единицы;Средняя Себестоимость 1 единицы', 'Кабель;Черный;A-01;Кабели;Поставщик 1;10;10;10;10;10;10;10;10;0;100;50', 'Кабель;Белый;A-02;Кабели;Поставщик 1;0;0;0;0;0;0;0;0;20;100;50', 'Наушники;Стандарт;B-01;Аудио;Поставщик 2;2;2;2;2;2;2;2;2;300;200;100', 'Чехол;Красный;C-01;Аксессуары;Поставщик 2;10;10;10;10;10;10;10;10;5;80;40'].join('\n');
async function call(path, { method = 'GET', body, cookie, csrf } = {}) { const response = await fetch(`${origin}${path}`, { method, headers: { ...(body ? { 'content-type': 'application/json' } : {}), ...(cookie ? { cookie } : {}), ...(csrf ? { 'x-csrf-token': csrf } : {}) }, body: body ? JSON.stringify(body) : undefined }); const text = await response.text(); return { response, data: text ? JSON.parse(text) : {}, cookie: response.headers.get('set-cookie')?.split(';')[0] }; }
try {
  const buyer = await store.createUser({ email: 'screens@example.test', name: 'Закупщик', password: 'BuyerPass!2026', role: 'BUYER' });
  const login = await call('/api/auth/login', { method: 'POST', body: { email: buyer.email, password: 'BuyerPass!2026' } });
  const preview = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'PASTE', text: source } } });
  const imported = await call('/api/imports/commit', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { previewId: preview.data.previewId, mapping: preview.data.suggested.mapping, periodColumns: preview.data.suggested.periodColumns, period: { granularity: 'week', lastPeriodEnd: '2026-08-24', order: 'OLDEST_TO_NEWEST' } } });
  assert.equal(imported.response.status, 201); assert.equal(imported.data.analysis.rows, 4);
  const summary = await call('/api/analytics/summary', { cookie: login.cookie }); assert.equal(summary.response.status, 200); assert.equal(summary.data.total, 4); assert.equal(summary.data.trend.length, 8); assert.ok(summary.data.matrix.length >= 2); assert.ok(summary.data.status.NO_STOCK >= 1); assert.ok(summary.data.orderNow.length >= 1); assert.ok(summary.data.totalRecommendedQty > 0); assert.ok(summary.data.topUrgent.length > 0); assert.ok(summary.data.topObsolete.length > 0);
  const all = await call('/api/analytics/results?pageSize=10', { cookie: login.cookie }); assert.equal(all.response.status, 200); assert.equal(all.data.total, 4); assert.equal(all.data.items.length, 4); assert.deepEqual(all.data.filters.categories, ['Аксессуары', 'Аудио', 'Кабели']); assert.deepEqual(all.data.filters.suppliers, ['Поставщик 1', 'Поставщик 2']); assert.ok(Array.isArray(all.data.items[0].warnings));
  const byCharacteristic = await call('/api/analytics/results?search=%D0%91%D0%B5%D0%BB%D1%8B%D0%B9', { cookie: login.cookie }); assert.equal(byCharacteristic.data.total, 1); assert.equal(byCharacteristic.data.items[0].characteristic_name, 'Белый');
  const byCode = await call('/api/analytics/results?search=B-01', { cookie: login.cookie }); assert.equal(byCode.data.total, 1); assert.equal(byCode.data.items[0].product_name, 'Наушники');
  const byStatus = await call('/api/analytics/results?status=NO_STOCK,ORDER_NOW&pageSize=10', { cookie: login.cookie }); assert.ok(byStatus.data.items.every((row) => ['NO_STOCK', 'ORDER_NOW'].includes(row.primary_status)));
  const page = await call('/api/analytics/results?page=2&pageSize=2&sort=name', { cookie: login.cookie }); assert.equal(page.data.page, 2); assert.equal(page.data.items.length, 2); assert.equal(page.data.total, 4);
  const detail = await call(`/api/analytics/products/${encodeURIComponent(all.data.items[0].product_id)}`, { cookie: login.cookie }); assert.equal(detail.response.status, 200); assert.equal(detail.data.periods.length, 8); assert.equal(detail.data.product.product_id, all.data.items[0].product_id); assert.ok(Number.isFinite(detail.data.formula.reorderPoint)); assert.ok(Number.isFinite(detail.data.formula.targetStock)); assert.ok(Array.isArray(detail.data.product.warnings));
  const missing = await call('/api/analytics/products/missing', { cookie: login.cookie }); assert.equal(missing.response.status, 404);
  console.log('Этап 4 аналитические API: 28 автоматических проверок пройдены.');
} finally { await new Promise((resolve) => server.close(resolve)); }
