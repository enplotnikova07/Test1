import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { makeApplication } from '../server.mjs';
import * as XLSX from 'xlsx';

const directory = await mkdtemp(join(tmpdir(), 'inventory-import-'));
const { server, store } = makeApplication({ databasePath: join(directory, 'test.db'), secureCookies: false });
await new Promise((resolve) => server.listen(0, resolve));
const origin = `http://127.0.0.1:${server.address().port}`;
const close = () => new Promise((resolve) => server.close(resolve));
const source = ['Номенклатура;Характеристика;1;2;3;4;5;6;7;8;Тек.остаток, шт;Средняя Цена 1 единицы;Средняя Себестоимость 1 единицы', 'Телефон;128 ГБ;2;3;1;4;2;1;3;2;10;50000;35000', 'Телефон;256 ГБ;1;2;2;3;1;2;2;1;5;60000;41000'].join('\n');
async function call(path, { method = 'GET', body, cookie, csrf } = {}) { const response = await fetch(`${origin}${path}`, { method, headers: { ...(body ? { 'content-type': 'application/json' } : {}), ...(cookie ? { cookie } : {}), ...(csrf ? { 'x-csrf-token': csrf } : {}) }, body: body ? JSON.stringify(body) : undefined }); const text = await response.text(); return { response, data: text ? JSON.parse(text) : {}, cookie: response.headers.get('set-cookie')?.split(';')[0] }; }

try {
  const buyer = await store.createUser({ email: 'buyer@example.test', name: 'Закупщик', password: 'BuyerPass!2026', role: 'BUYER' });
  const admin = await store.createUser({ email: 'admin@example.test', name: 'Админ', password: 'AdminPass!2026', role: 'ADMIN' });
  const login = await call('/api/auth/login', { method: 'POST', body: { email: buyer.email, password: 'BuyerPass!2026' } });
  const preview = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'PASTE', text: source } } });
  assert.equal(preview.response.status, 200); assert.equal(preview.data.rowCount, 2); assert.deepEqual(preview.data.suggested.periodColumns, ['1', '2', '3', '4', '5', '6', '7', '8']); assert.equal(preview.data.suggested.mapping.productName, 'Номенклатура');
  const commitInput = { previewId: preview.data.previewId, mapping: preview.data.suggested.mapping, periodColumns: preview.data.suggested.periodColumns, period: { granularity: 'week', lastPeriodEnd: '2026-08-24', order: 'OLDEST_TO_NEWEST' }, financialType: 'NONE', profileAction: 'CREATE', profileName: 'Стандартная выгрузка 1С' };
  const committed = await call('/api/imports/commit', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: commitInput });
  assert.equal(committed.response.status, 201); assert.equal(committed.data.rowCount, 2); const firstVersion = committed.data.versionId;
  assert.equal(store.activeVersion().id, firstVersion); assert.equal(store.db.prepare('SELECT count(*) AS count FROM version_products WHERE version_id=?').get(firstVersion).count, 2); assert.equal(store.db.prepare('SELECT count(*) AS count FROM sales_periods WHERE version_id=?').get(firstVersion).count, 16);
  const profiles = await call('/api/import-profiles', { cookie: login.cookie }); assert.equal(profiles.response.status, 200); assert.equal(profiles.data.profiles.length, 1);
  const xlsxBook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(xlsxBook, XLSX.utils.aoa_to_sheet([['служебный лист'], ['не импортировать']]), 'Служебный'); XLSX.utils.book_append_sheet(xlsxBook, XLSX.utils.aoa_to_sheet(source.split('\n').map((line) => line.split(';'))), 'Продажи');
  const xlsxBase64 = XLSX.write(xlsxBook, { type: 'base64', bookType: 'xlsx' });
  const xlsxPreview = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'FILE', fileName: 'выгрузка.xlsx', contentBase64: xlsxBase64 } } }); assert.equal(xlsxPreview.response.status, 200); assert.deepEqual(xlsxPreview.data.sheets, ['Служебный', 'Продажи']);
  const secondSheet = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'FILE', fileName: 'выгрузка.xlsx', contentBase64: xlsxBase64, sheetName: 'Продажи' } } }); assert.equal(secondSheet.data.rowCount, 2); assert.equal(secondSheet.data.sheetName, 'Продажи');
  const csvBase64 = Buffer.from(source, 'utf8').toString('base64'); const csvPreview = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'FILE', fileName: 'выгрузка.csv', contentBase64: csvBase64 } } }); assert.equal(csvPreview.data.rowCount, 2);
  const demoPreview = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'DEMO' } } }); assert.equal(demoPreview.data.rowCount, 60);
  const tooLarge = ['Номенклатура;1;Тек.остаток, шт', ...Array.from({ length: 50001 }, (_, index) => `Товар ${index};1;1`)].join('\n');
  const rejectedLarge = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'PASTE', text: tooLarge } } }); assert.equal(rejectedLarge.response.status, 400);
  const matchingPreview = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'PASTE', text: source } } }); assert.equal(matchingPreview.data.matchedProfiles.length, 1); assert.equal(matchingPreview.data.suggested.mapping.productName, 'Номенклатура');
  const archivedProfile = await call(`/api/import-profiles/${profiles.data.profiles[0].id}/archive`, { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken }); assert.equal(archivedProfile.response.status, 200);
  const noProfilePreview = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'PASTE', text: source } } }); assert.equal(noProfilePreview.data.matchedProfiles.length, 0);
  const duplicate = `${source}\nТелефон;128 ГБ;1;1;1;1;1;1;1;1;3;50000;35000`;
  const badPreview = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'PASTE', text: duplicate } } });
  const badCommit = await call('/api/imports/commit', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { previewId: badPreview.data.previewId, mapping: badPreview.data.suggested.mapping, periodColumns: badPreview.data.suggested.periodColumns, period: { granularity: 'week', lastPeriodEnd: '2026-08-24', order: 'OLDEST_TO_NEWEST' } } }); assert.equal(badCommit.response.status, 422); assert.equal(store.activeVersion().id, firstVersion);
  const changed = source.replace(';10;50000;35000', ';99;50000;35000');
  const changedPreview = await call('/api/imports/preview', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { source: { type: 'PASTE', text: changed } } });
  const changedCommit = await call('/api/imports/commit', { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken, body: { previewId: changedPreview.data.previewId, mapping: changedPreview.data.suggested.mapping, periodColumns: changedPreview.data.suggested.periodColumns, period: { granularity: 'week', lastPeriodEnd: '2026-08-24', order: 'OLDEST_TO_NEWEST' } } }); assert.equal(changedCommit.response.status, 201); assert.notEqual(changedCommit.data.versionId, firstVersion);
  const restored = await call(`/api/versions/${firstVersion}/restore`, { method: 'POST', cookie: login.cookie, csrf: login.data.csrfToken }); assert.equal(restored.response.status, 201); assert.notEqual(restored.data.versionId, firstVersion); assert.equal(store.activeVersion().id, restored.data.versionId); assert.equal(store.db.prepare('SELECT stock_qty FROM version_products WHERE version_id=? AND characteristic_name=?').get(restored.data.versionId, '128 ГБ').stock_qty, 10);
  const adminLogin = await call('/api/auth/login', { method: 'POST', body: { email: admin.email, password: 'AdminPass!2026' } }); const forbidden = await call('/api/imports/preview', { method: 'POST', cookie: adminLogin.cookie, csrf: adminLogin.data.csrfToken, body: { source: { type: 'PASTE', text: source } } }); assert.equal(forbidden.response.status, 403);
  console.log('Этап 2 сервер: 37 автоматических проверок пройдены.');
} finally { await close(); }
