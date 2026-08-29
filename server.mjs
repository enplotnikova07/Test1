import { createHash, randomBytes, scrypt as scryptCallback, timingSafeEqual } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { mkdirSync } from 'node:fs';
import { createServer as createHttpServer } from 'node:http';
import { connect as connectNet } from 'node:net';
import { connect as connectTls } from 'node:tls';
import { dirname, extname, join, normalize } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { DatabaseSync } from 'node:sqlite';
import * as XLSX from 'xlsx';

const scrypt = promisify(scryptCallback);
const here = dirname(fileURLToPath(import.meta.url));
const roles = new Set(['BUYER', 'MANAGER', 'FINANCE', 'ADMIN']);
const roleLabels = { BUYER: 'Закупщик', MANAGER: 'Руководитель', FINANCE: 'Финансовый директор', ADMIN: 'Администратор' };
const mimeTypes = { '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.svg': 'image/svg+xml' };
const now = () => new Date().toISOString();
const id = () => randomBytes(18).toString('base64url');
const digest = (value) => createHash('sha256').update(value).digest('hex');
const importFields = [
  ['productName', 'Номенклатура', true], ['characteristic', 'Характеристика', false], ['characteristicCode', 'Код характеристики', false], ['externalCode', 'Код 1С', false],
  ['category', 'Категория', false], ['supplier', 'Поставщик', false], ['stockQty', 'Текущий остаток, шт', true], ['stockValue', 'Текущий остаток, руб', false],
  ['reserveQty', 'В резерве', false], ['inTransitQty', 'В пути', false], ['avgCost', 'Средняя себестоимость', false], ['avgPrice', 'Средняя цена', false],
  ['leadTimeDays', 'Срок поставки, дней', false], ['delayDays', 'Задержка поставки, дней', false], ['reviewPeriodDays', 'Интервал заказа, дней', false],
  ['sourceFinancialTotal', 'Финансовый итог', false], ['sourceTotalSold', 'Всего продано, шт', false], ['monthsWithSales', 'Месяцев продаж', false]
];
const aliases = {
  productName: ['номенклатура', 'наименование', 'товар'], characteristic: ['характеристика номенклатуры', 'характеристика'], characteristicCode: ['код характеристики'], externalCode: ['код 1с', 'код товара'],
  category: ['категория'], supplier: ['поставщик'], stockQty: ['тек.остаток, шт', 'текущий остаток, шт', 'остаток, шт', 'остаток'], stockValue: ['тек.остаток, руб', 'текущий остаток, руб', 'остаток, руб'],
  reserveQty: ['в резерве', 'резерв'], inTransitQty: ['в пути'], avgCost: ['средняя себестоимость 1 единицы', 'средняя себестоимость'], avgPrice: ['средняя цена 1 единицы', 'средняя цена'],
  leadTimeDays: ['за сколько нам приходит товар от поставщика, дн', 'срок поставки, дней'], delayDays: ['на сколько по времени могут задержать поставки, дни', 'задержка поставки, дней'],
  reviewPeriodDays: ['как часто мы заказываем товар, дни', 'интервал заказа, дней'], sourceFinancialTotal: ['объем продаж/маржа (наценка) за весь период, руб', 'выручка', 'валовая прибыль'],
  sourceTotalSold: ['всего продано, шт'], monthsWithSales: ['кол-во месяцев продажи']
};
const normalizeText = (value) => String(value ?? '').trim().replace(/\s+/g, ' ');
const normalizeKey = (value) => normalizeText(value).toLocaleLowerCase('ru-RU');
const defaultSettings = Object.freeze({ abcA: 80, abcB: 95, xyzX: 10, xyzY: 25, minHistory: 8, defaultLeadDays: 5, defaultDelayDays: 0, defaultReviewDays: 7, earlyWarningDays: 3, slowDays: 60, obsoleteDays: 90, excessDays: 90, safetyZ: 1.65, excludeOutliers: false });
const numberOr = (value, fallback) => Number.isFinite(Number(value)) ? Number(value) : fallback;
function median(values) { if (!values.length) return 0; const sorted = [...values].sort((a, b) => a - b); const middle = Math.floor(sorted.length / 2); return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2; }
function percentile(sorted, point) { if (!sorted.length) return 0; const index = (sorted.length - 1) * point; const lower = Math.floor(index); const upper = Math.ceil(index); return lower === upper ? sorted[lower] : sorted[lower] + (sorted[upper] - sorted[lower]) * (index - lower); }
function average(values) { return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0; }
function stddevPopulation(values, mean = average(values)) { return values.length ? Math.sqrt(values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / values.length) : 0; }
function numberValue(value) {
  if (value === null || value === undefined || value === '') return { value: null, valid: true };
  if (typeof value === 'number' && Number.isFinite(value)) return { value, valid: true };
  const parsed = Number(String(value).replace(/\s/g, '').replace(',', '.'));
  return Number.isFinite(parsed) ? { value: parsed, valid: true } : { value: null, valid: false };
}
function defaultMapping(headers) {
  const normalized = headers.map(normalizeKey);
  const mapping = {};
  for (const [field] of importFields) {
    const index = normalized.findIndex((header) => (aliases[field] || []).includes(header));
    if (index >= 0) mapping[field] = headers[index];
  }
  return { mapping, periodColumns: headers.filter((header) => /^\d+$/.test(normalizeText(header))) };
}
function uniqueHeaders(input) {
  const seen = new Map();
  return input.map((value, index) => { const base = normalizeText(value) || `Колонка ${index + 1}`; const count = (seen.get(base) || 0) + 1; seen.set(base, count); return count === 1 ? base : `${base} (${count})`; });
}
function tableFromWorksheet(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
  const headerIndex = Math.max(0, matrix.findIndex((row) => Array.isArray(row) && row.some((value) => ['номенклатура', 'наименование'].includes(normalizeKey(value)))));
  const headers = uniqueHeaders(matrix[headerIndex] || []);
  const rows = matrix.slice(headerIndex + 1).filter((row) => row.some((value) => value !== null && normalizeText(value) !== '')).map((row, offset) => ({ rowNumber: headerIndex + offset + 2, values: Object.fromEntries(headers.map((header, index) => [header, row[index] ?? null])) }));
  return { headers, rows, headerRow: headerIndex + 1 };
}
function parseWorkbook(source) {
  let workbook;
  if (source.kind === 'text') workbook = XLSX.read(source.content, { type: 'string', raw: true });
  else workbook = XLSX.read(Buffer.from(source.content, 'base64'), { type: 'buffer', raw: true });
  const sheetName = source.sheetName && workbook.SheetNames.includes(source.sheetName) ? source.sheetName : workbook.SheetNames[0];
  if (!sheetName) throw new Error('В файле нет листов с данными.');
  const table = tableFromWorksheet(workbook.Sheets[sheetName]);
  return { ...table, sheetName, sheets: workbook.SheetNames };
}
function parsePeriodDate(lastPeriodEnd, granularity, index, total, order) {
  const end = new Date(`${lastPeriodEnd}T00:00:00.000Z`);
  const position = order === 'NEWEST_TO_OLDEST' ? index : total - index - 1;
  if (granularity === 'day') end.setUTCDate(end.getUTCDate() - position);
  else if (granularity === 'week') end.setUTCDate(end.getUTCDate() - position * 7);
  else end.setUTCMonth(end.getUTCMonth() - position);
  const start = new Date(end);
  if (granularity === 'week') start.setUTCDate(start.getUTCDate() - 6);
  if (granularity === 'month') start.setUTCDate(1);
  return { start: start.toISOString().slice(0, 10), end: end.toISOString().slice(0, 10) };
}
function normalizeImport(table, { mapping, periodColumns, period, financialType = 'NONE' }) {
  const issues = []; const rows = []; const usedPeriods = (periodColumns || []).filter((column) => table.headers.includes(column));
  if (!mapping?.productName) issues.push({ severity: 'ERROR', code: 'PRODUCT_REQUIRED', message: 'Не сопоставлена обязательная колонка «Номенклатура».', column: null });
  if (!mapping?.stockQty) issues.push({ severity: 'ERROR', code: 'STOCK_REQUIRED', message: 'Не сопоставлена обязательная колонка «Текущий остаток, шт».', column: null });
  if (!usedPeriods.length) issues.push({ severity: 'ERROR', code: 'PERIODS_REQUIRED', message: 'Не выбраны колонки продаж за периоды.', column: null });
  if (!['day', 'week', 'month'].includes(period?.granularity)) issues.push({ severity: 'ERROR', code: 'PERIOD_GRANULARITY_REQUIRED', message: 'Выберите шаг периодов: день, неделя или месяц.', column: null });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(period?.lastPeriodEnd || '')) issues.push({ severity: 'ERROR', code: 'LAST_PERIOD_REQUIRED', message: 'Укажите дату окончания последнего периода.', column: null });
  if (issues.some((issue) => issue.severity === 'ERROR')) return { rows, issues, hasErrors: true };
  const headerValue = (values, field) => mapping[field] ? values[mapping[field]] : null;
  const numericField = (values, field, rowNo, required = false) => {
    const label = mapping[field]; const parsed = numberValue(headerValue(values, field));
    if (required && parsed.value === null) issues.push({ severity: 'ERROR', code: 'NUMBER_REQUIRED', message: `В обязательной колонке «${label}» нет числового значения.`, rowNo, column: label });
    else if (!parsed.valid) issues.push({ severity: 'ERROR', code: 'NUMBER_INVALID', message: `Значение в колонке «${label}» не является числом.`, rowNo, column: label });
    return parsed.value;
  };
  const keys = new Set();
  for (const input of table.rows) {
    const productName = normalizeText(headerValue(input.values, 'productName')); const characteristicName = normalizeText(headerValue(input.values, 'characteristic')); const externalCode = normalizeText(headerValue(input.values, 'externalCode')) || null; const characteristicCode = normalizeText(headerValue(input.values, 'characteristicCode')) || null;
    if (!productName) { issues.push({ severity: 'ERROR', code: 'PRODUCT_EMPTY', message: 'Не заполнена номенклатура.', rowNo: input.rowNumber, column: mapping.productName }); continue; }
    const normalizedName = normalizeKey(productName); const normalizedCharacteristicName = normalizeKey(characteristicName); const compositeKey = externalCode ? `${externalCode}::${characteristicCode || normalizedCharacteristicName}` : `${normalizedName}::${normalizedCharacteristicName}`;
    if (keys.has(compositeKey)) { issues.push({ severity: 'ERROR', code: 'DUPLICATE_PRODUCT', message: 'Повторяется составной ключ номенклатуры и характеристики.', rowNo: input.rowNumber, column: mapping.productName }); continue; }
    keys.add(compositeKey);
    const periods = usedPeriods.map((column, index) => { const parsed = numberValue(input.values[column]); if (!parsed.valid) issues.push({ severity: 'ERROR', code: 'PERIOD_INVALID', message: `Продажи в периоде «${column}» не являются числом.`, rowNo: input.rowNumber, column }); const value = parsed.value ?? 0; if (value < 0) issues.push({ severity: 'WARNING', code: 'NEGATIVE_SALES', message: `Нетто-продажа в периоде «${column}» отрицательна: для спроса будет использован 0.`, rowNo: input.rowNumber, column }); const dates = parsePeriodDate(period.lastPeriodEnd, period.granularity, index, usedPeriods.length, period.order); return { label: column, ...dates, value }; });
    const stockQty = numericField(input.values, 'stockQty', input.rowNumber, true); const sourceTotalSold = numericField(input.values, 'sourceTotalSold', input.rowNumber); const stockValue = numericField(input.values, 'stockValue', input.rowNumber); const reserveQty = numericField(input.values, 'reserveQty', input.rowNumber); const inTransitQty = numericField(input.values, 'inTransitQty', input.rowNumber); const avgCost = numericField(input.values, 'avgCost', input.rowNumber); const avgPrice = numericField(input.values, 'avgPrice', input.rowNumber); const leadTimeDays = numericField(input.values, 'leadTimeDays', input.rowNumber); const delayDays = numericField(input.values, 'delayDays', input.rowNumber); const reviewPeriodDays = numericField(input.values, 'reviewPeriodDays', input.rowNumber); const sourceFinancialTotal = numericField(input.values, 'sourceFinancialTotal', input.rowNumber); const monthsWithSales = numericField(input.values, 'monthsWithSales', input.rowNumber);
    if (stockQty !== null && stockQty < 0) issues.push({ severity: 'WARNING', code: 'NEGATIVE_STOCK', message: 'Текущий остаток отрицательный.', rowNo: input.rowNumber, column: mapping.stockQty });
    if (avgCost === null || avgPrice === null) issues.push({ severity: 'WARNING', code: 'PRICE_OR_COST_MISSING', message: 'Не указана средняя цена или себестоимость.', rowNo: input.rowNumber });
    if (leadTimeDays === null) issues.push({ severity: 'WARNING', code: 'LEAD_TIME_MISSING', message: 'Не указан срок поставки: будет использовано значение из будущих настроек.', rowNo: input.rowNumber });
    const periodSum = periods.reduce((sum, item) => sum + Math.max(0, item.value), 0);
    if (sourceTotalSold !== null && Math.abs(sourceTotalSold - periodSum) > Math.max(1, Math.abs(sourceTotalSold) * 0.01)) issues.push({ severity: 'WARNING', code: 'TOTAL_SOLD_MISMATCH', message: 'Итог продаж отличается от суммы периодов более чем на 1%.', rowNo: input.rowNumber, column: mapping.sourceTotalSold });
    if (periods.length < 8) issues.push({ severity: 'WARNING', code: 'SHORT_HISTORY', message: 'История содержит менее 8 периодов.', rowNo: input.rowNumber });
    rows.push({ externalCode, productName, normalizedName, characteristicCode, characteristicName: characteristicName || null, normalizedCharacteristicName, compositeKey, category: normalizeText(headerValue(input.values, 'category')) || null, supplier: normalizeText(headerValue(input.values, 'supplier')) || null, stockQty, stockValue, reserveQty, inTransitQty, avgCost, avgPrice, leadTimeDays, delayDays, reviewPeriodDays, sourceFinancialTotal, sourceTotalSold, monthsWithSales, periods });
  }
  return { rows, issues, hasErrors: issues.some((issue) => issue.severity === 'ERROR') };
}
function exampleCsv() {
  const header = ['Номенклатура', 'Характеристика', '1', '2', '3', '4', '5', '6', 'Тек.остаток, шт', 'Средняя Себестоимость 1 единицы', 'Средняя Цена 1 единицы', 'За сколько нам приходит товар от поставщика, дн'];
  const rows = Array.from({ length: 60 }, (_, index) => [`Демо-товар ${String(index + 1).padStart(2, '0')}`, index % 3 === 0 ? `Модификация ${index % 5 + 1}` : '', 2 + index % 7, 3 + index % 5, index % 4, 4 + index % 8, 2 + index % 6, 5 + index % 9, 5 + index % 30, 400 + index * 15, 700 + index * 20, 5]).map((row) => row.join(';'));
  return [header.join(';'), ...rows].join('\n');
}

async function passwordHash(password) {
  if (typeof password !== 'string' || password.length < 10) throw new Error('Пароль должен содержать не менее 10 символов.');
  const salt = randomBytes(16).toString('base64url');
  const hash = await scrypt(password, salt, 64);
  return `scrypt$${salt}$${Buffer.from(hash).toString('base64url')}`;
}
async function passwordMatches(password, stored) {
  const [kind, salt, expected] = String(stored).split('$');
  if (kind !== 'scrypt' || !salt || !expected) return false;
  const actual = Buffer.from(await scrypt(password, salt, 64));
  const target = Buffer.from(expected, 'base64url');
  return actual.length === target.length && timingSafeEqual(actual, target);
}
function safeUser(row) {
  return { id: row.id, email: row.email, name: row.name, role: row.role, roleLabel: roleLabels[row.role], active: Boolean(row.active), createdAt: row.created_at, lastLoginAt: row.last_login_at };
}

export class Store {
  constructor(path) {
    mkdirSync(dirname(path), { recursive: true });
    this.db = new DatabaseSync(path);
    this.db.exec('PRAGMA journal_mode = WAL; PRAGMA foreign_keys = ON;');
    this.migrate();
  }
  migrate() {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS users (
        id TEXT PRIMARY KEY, email TEXT NOT NULL UNIQUE COLLATE NOCASE, name TEXT NOT NULL,
        password_hash TEXT NOT NULL, role TEXT NOT NULL CHECK(role IN ('BUYER','MANAGER','FINANCE','ADMIN')),
        active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL, last_login_at TEXT
      );
      CREATE TABLE IF NOT EXISTS sessions (
        id TEXT PRIMARY KEY, user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        csrf_token TEXT NOT NULL, expires_at TEXT NOT NULL, created_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS password_resets (
        token_hash TEXT PRIMARY KEY, user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        expires_at TEXT NOT NULL, used_at TEXT, created_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS audit_events (
        id TEXT PRIMARY KEY, actor_id TEXT REFERENCES users(id), action TEXT NOT NULL,
        entity_type TEXT NOT NULL, entity_id TEXT, metadata_json TEXT NOT NULL DEFAULT '{}', created_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS data_versions (
        id TEXT PRIMARY KEY, status TEXT NOT NULL CHECK(status IN ('ACTIVE','ARCHIVED','FAILED')), source_type TEXT NOT NULL,
        source_name TEXT NOT NULL, source_hash TEXT NOT NULL, created_by TEXT NOT NULL REFERENCES users(id), created_at TEXT NOT NULL,
        activated_at TEXT, parent_version_id TEXT REFERENCES data_versions(id), period_granularity TEXT NOT NULL,
        last_period_end TEXT NOT NULL, periods_order TEXT NOT NULL, row_count INTEGER NOT NULL, product_count INTEGER NOT NULL,
        profile_id TEXT, financial_type TEXT NOT NULL DEFAULT 'NONE'
      );
      CREATE TABLE IF NOT EXISTS version_products (
        id TEXT PRIMARY KEY, version_id TEXT NOT NULL REFERENCES data_versions(id) ON DELETE CASCADE, external_code TEXT,
        product_name TEXT NOT NULL, normalized_name TEXT NOT NULL, characteristic_code TEXT, characteristic_name TEXT,
        normalized_characteristic_name TEXT NOT NULL, composite_key TEXT NOT NULL, category TEXT, supplier TEXT,
        stock_qty REAL, stock_value REAL, reserve_qty REAL, in_transit_qty REAL, avg_cost REAL, avg_price REAL,
        lead_time_days REAL, delay_days REAL, review_period_days REAL, source_financial_total REAL, source_total_sold REAL, months_with_sales REAL
      );
      CREATE TABLE IF NOT EXISTS sales_periods (
        id TEXT PRIMARY KEY, version_id TEXT NOT NULL REFERENCES data_versions(id) ON DELETE CASCADE,
        product_id TEXT NOT NULL REFERENCES version_products(id) ON DELETE CASCADE, period_number INTEGER NOT NULL,
        period_label TEXT NOT NULL, period_start TEXT, period_end TEXT, net_qty REAL NOT NULL
      );
      CREATE TABLE IF NOT EXISTS imports_issues (
        id TEXT PRIMARY KEY, version_id TEXT NOT NULL REFERENCES data_versions(id) ON DELETE CASCADE, row_no INTEGER,
        column_name TEXT, severity TEXT NOT NULL CHECK(severity IN ('ERROR','WARNING')), code TEXT NOT NULL, message TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS import_profiles (
        id TEXT PRIMARY KEY, name TEXT NOT NULL, status TEXT NOT NULL CHECK(status IN ('ACTIVE','ARCHIVED')), source_type TEXT NOT NULL,
        columns_signature TEXT NOT NULL, mapping_json TEXT NOT NULL, period_columns_json TEXT NOT NULL, period_granularity TEXT,
        periods_order TEXT, financial_type TEXT NOT NULL DEFAULT 'NONE', version_no INTEGER NOT NULL DEFAULT 1,
        created_by TEXT NOT NULL REFERENCES users(id), created_at TEXT NOT NULL, updated_at TEXT NOT NULL, archived_at TEXT
      );
      CREATE TABLE IF NOT EXISTS calculation_settings (
        id TEXT PRIMARY KEY, params_json TEXT NOT NULL, created_by TEXT REFERENCES users(id), created_at TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1
      );
      CREATE TABLE IF NOT EXISTS analysis_results (
        id TEXT PRIMARY KEY, version_id TEXT NOT NULL REFERENCES data_versions(id) ON DELETE CASCADE,
        settings_id TEXT NOT NULL REFERENCES calculation_settings(id), product_id TEXT NOT NULL REFERENCES version_products(id) ON DELETE CASCADE,
        total_sold REAL NOT NULL, revenue REAL NOT NULL, gross_profit REAL NOT NULL, abc_revenue TEXT NOT NULL, abc_profit TEXT NOT NULL, abc_quantity TEXT NOT NULL,
        xyz TEXT, cv REAL, average_demand REAL NOT NULL, demand_stddev REAL NOT NULL, demand_change_pct REAL, outlier_count INTEGER NOT NULL,
        free_stock REAL NOT NULL, stock_position REAL NOT NULL, effective_lead_days REAL NOT NULL, safety_stock REAL NOT NULL, reorder_point REAL NOT NULL,
        target_stock REAL NOT NULL, recommended_qty INTEGER NOT NULL, days_of_stock REAL, last_sale_date TEXT, primary_status TEXT NOT NULL,
        warnings_json TEXT NOT NULL DEFAULT '[]', calculated_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS manual_adjustments (
        id TEXT PRIMARY KEY, version_id TEXT NOT NULL REFERENCES data_versions(id), product_id TEXT NOT NULL REFERENCES version_products(id),
        calculated_qty INTEGER NOT NULL, adjusted_qty INTEGER NOT NULL, reason TEXT NOT NULL, created_by TEXT NOT NULL REFERENCES users(id), created_at TEXT NOT NULL, reverted_at TEXT, reverted_by TEXT REFERENCES users(id)
      );
      CREATE TABLE IF NOT EXISTS orders (
        id TEXT PRIMARY KEY, number TEXT NOT NULL UNIQUE, title TEXT NOT NULL, status TEXT NOT NULL CHECK(status IN ('DRAFT','FORMED','ARCHIVED')),
        supplier TEXT, comment TEXT, data_version_id TEXT NOT NULL REFERENCES data_versions(id), settings_version_id TEXT NOT NULL REFERENCES calculation_settings(id), created_by TEXT NOT NULL REFERENCES users(id), created_at TEXT NOT NULL, formed_at TEXT, archived_at TEXT
      );
      CREATE TABLE IF NOT EXISTS order_items (
        id TEXT PRIMARY KEY, order_id TEXT NOT NULL REFERENCES orders(id) ON DELETE CASCADE, product_id TEXT NOT NULL REFERENCES version_products(id), calculated_qty INTEGER NOT NULL, adjusted_qty INTEGER, final_qty INTEGER NOT NULL, unit_cost REAL, snapshot_json TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS order_status_history (id TEXT PRIMARY KEY, order_id TEXT NOT NULL REFERENCES orders(id) ON DELETE CASCADE, from_status TEXT, to_status TEXT NOT NULL, actor_id TEXT NOT NULL REFERENCES users(id), comment TEXT, created_at TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS email_deliveries (id TEXT PRIMARY KEY, report_type TEXT NOT NULL, recipients_json TEXT NOT NULL, subject TEXT, message TEXT, status TEXT NOT NULL, error TEXT, created_by TEXT NOT NULL REFERENCES users(id), created_at TEXT NOT NULL, sent_at TEXT);
      CREATE INDEX IF NOT EXISTS idx_versions_active ON data_versions(status, activated_at DESC);
      CREATE INDEX IF NOT EXISTS idx_products_version ON version_products(version_id);
      CREATE INDEX IF NOT EXISTS idx_periods_product ON sales_periods(product_id, period_number);
      CREATE INDEX IF NOT EXISTS idx_profiles_signature ON import_profiles(columns_signature, status);
      CREATE INDEX IF NOT EXISTS idx_results_version ON analysis_results(version_id, settings_id);
      CREATE INDEX IF NOT EXISTS idx_adjustments_active ON manual_adjustments(version_id, product_id, reverted_at);
      CREATE INDEX IF NOT EXISTS idx_orders_status ON orders(status, created_at DESC);
    `);
    if (!this.db.prepare('SELECT id FROM calculation_settings WHERE active=1 LIMIT 1').get()) this.db.prepare('INSERT INTO calculation_settings (id,params_json,created_by,created_at,active) VALUES (?,?,?,?,1)').run(id(), JSON.stringify(defaultSettings), null, now());
  }
  async createUser({ email, name, password, role }, actorId = null) {
    const normalizedEmail = String(email || '').trim().toLowerCase();
    const displayName = String(name || '').trim();
    if (!/^\S+@\S+\.\S+$/.test(normalizedEmail)) throw new Error('Укажите корректный email.');
    if (!displayName) throw new Error('Укажите имя пользователя.');
    if (!roles.has(role)) throw new Error('Укажите корректную роль.');
    const user = { id: id(), email: normalizedEmail, name: displayName, passwordHash: await passwordHash(password), role, createdAt: now() };
    try { this.db.prepare('INSERT INTO users (id,email,name,password_hash,role,active,created_at) VALUES (?,?,?,?,?,1,?)').run(user.id, user.email, user.name, user.passwordHash, user.role, user.createdAt); }
    catch (error) { if (String(error.message).includes('UNIQUE')) throw new Error('Пользователь с таким email уже существует.'); throw error; }
    this.audit(actorId, 'USER_CREATED', 'USER', user.id, { email: user.email, role: user.role });
    return safeUser(this.userById(user.id));
  }
  userByEmail(email) { return this.db.prepare('SELECT * FROM users WHERE email = ?').get(String(email || '').trim().toLowerCase()); }
  userById(userId) { return this.db.prepare('SELECT * FROM users WHERE id = ?').get(userId); }
  listUsers() { return this.db.prepare('SELECT * FROM users ORDER BY created_at DESC').all().map(safeUser); }
  activeAdminCount() { return this.db.prepare("SELECT count(*) AS count FROM users WHERE role = 'ADMIN' AND active = 1").get().count; }
  updateUser(userId, { name, role }, actorId) {
    const existing = this.userById(userId); if (!existing) throw new Error('Пользователь не найден.');
    const nextName = name === undefined ? existing.name : String(name).trim(); const nextRole = role === undefined ? existing.role : role;
    if (!nextName) throw new Error('Укажите имя пользователя.'); if (!roles.has(nextRole)) throw new Error('Укажите корректную роль.');
    if (existing.role === 'ADMIN' && existing.active && nextRole !== 'ADMIN' && this.activeAdminCount() <= 1) throw new Error('Нельзя снять роль с последнего активного администратора.');
    this.db.prepare('UPDATE users SET name = ?, role = ? WHERE id = ?').run(nextName, nextRole, userId);
    this.audit(actorId, 'USER_UPDATED', 'USER', userId, { name: nextName, role: nextRole }); return safeUser(this.userById(userId));
  }
  setUserActive(userId, active, actorId) {
    const existing = this.userById(userId); if (!existing) throw new Error('Пользователь не найден.');
    if (!active && existing.role === 'ADMIN' && existing.active && this.activeAdminCount() <= 1) throw new Error('Нельзя заблокировать последнего активного администратора.');
    this.db.prepare('UPDATE users SET active = ? WHERE id = ?').run(active ? 1 : 0, userId);
    if (!active) this.db.prepare('DELETE FROM sessions WHERE user_id = ?').run(userId);
    this.audit(actorId, active ? 'USER_UNBLOCKED' : 'USER_BLOCKED', 'USER', userId, {}); return safeUser(this.userById(userId));
  }
  createReset(userId, actorId) {
    if (!this.userById(userId)) throw new Error('Пользователь не найден.'); const token = id(); const expiresAt = new Date(Date.now() + 30 * 60 * 1000).toISOString();
    this.db.prepare('INSERT INTO password_resets (token_hash,user_id,expires_at,created_at) VALUES (?,?,?,?)').run(createHash('sha256').update(token).digest('hex'), userId, expiresAt, now());
    this.audit(actorId, 'PASSWORD_RESET_REQUESTED', 'USER', userId, { expiresAt }); return { token, expiresAt };
  }
  async resetPassword(token, password) {
    const digest = createHash('sha256').update(token).digest('hex'); const reset = this.db.prepare('SELECT * FROM password_resets WHERE token_hash = ?').get(digest);
    if (!reset || reset.used_at || new Date(reset.expires_at) < new Date()) throw new Error('Ссылка сброса пароля недействительна или истекла.');
    this.db.prepare('UPDATE users SET password_hash = ? WHERE id = ?').run(await passwordHash(password), reset.user_id);
    this.db.prepare('UPDATE password_resets SET used_at = ? WHERE token_hash = ?').run(now(), digest); this.db.prepare('DELETE FROM sessions WHERE user_id = ?').run(reset.user_id); this.audit(null, 'PASSWORD_RESET_COMPLETED', 'USER', reset.user_id, {});
  }
  createSession(userId) { const session = { id: id(), csrf: id(), expires: new Date(Date.now() + 8 * 60 * 60 * 1000).toISOString() }; this.db.prepare('INSERT INTO sessions (id,user_id,csrf_token,expires_at,created_at) VALUES (?,?,?,?,?)').run(session.id, userId, session.csrf, session.expires, now()); return session; }
  session(sessionId) { if (!sessionId) return null; const row = this.db.prepare('SELECT s.id AS session_id,s.csrf_token,s.expires_at,u.* FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.id=?').get(sessionId); if (!row || !row.active || new Date(row.expires_at) < new Date()) { if (row) this.db.prepare('DELETE FROM sessions WHERE id = ?').run(sessionId); return null; } return { user: safeUser(row), csrf: row.csrf_token }; }
  deleteSession(sessionId) { if (sessionId) this.db.prepare('DELETE FROM sessions WHERE id = ?').run(sessionId); }
  setLastLogin(userId) { this.db.prepare('UPDATE users SET last_login_at = ? WHERE id = ?').run(now(), userId); }
  audit(actorId, action, entityType, entityId, metadata) { this.db.prepare('INSERT INTO audit_events (id,actor_id,action,entity_type,entity_id,metadata_json,created_at) VALUES (?,?,?,?,?,?,?)').run(id(), actorId, action, entityType, entityId, JSON.stringify(metadata || {}), now()); }
  listAudit() { return this.db.prepare('SELECT a.*, u.email AS actor_email FROM audit_events a LEFT JOIN users u ON u.id=a.actor_id ORDER BY a.created_at DESC LIMIT 200').all(); }
  activeSettings() { const item = this.db.prepare('SELECT * FROM calculation_settings WHERE active=1 ORDER BY created_at DESC LIMIT 1').get(); return { id: item.id, params: { ...defaultSettings, ...JSON.parse(item.params_json) }, createdAt: item.created_at }; }
  settingsHistory() { return this.db.prepare('SELECT id,params_json,created_at,active FROM calculation_settings ORDER BY created_at DESC LIMIT 20').all().map((item) => ({ id: item.id, params: { ...defaultSettings, ...JSON.parse(item.params_json) }, createdAt: item.created_at, active: Boolean(item.active) })); }
  updateSettings(params, actorId) {
    const next = { ...defaultSettings, ...params }; const numberKeys = ['abcA','abcB','xyzX','xyzY','minHistory','defaultLeadDays','defaultDelayDays','defaultReviewDays','earlyWarningDays','slowDays','obsoleteDays','excessDays','safetyZ'];
    for (const key of numberKeys) next[key] = numberOr(next[key], defaultSettings[key]);
    if (!(next.abcA > 0 && next.abcB > next.abcA && next.abcB <= 100)) throw new Error('Границы ABC должны соответствовать правилу 0 < A < B ≤ 100.');
    if (!(next.xyzX >= 0 && next.xyzY > next.xyzX)) throw new Error('Границы XYZ должны соответствовать правилу 0 ≤ X < Y.');
    if (!(next.minHistory >= 3 && next.defaultLeadDays >= 0 && next.defaultReviewDays >= 1 && next.obsoleteDays > next.slowDays)) throw new Error('Проверьте минимальную историю, сроки и пороги статусов.');
    const settingsId = id(); this.db.exec('BEGIN IMMEDIATE');
    try { this.db.prepare('UPDATE calculation_settings SET active=0 WHERE active=1').run(); this.db.prepare('INSERT INTO calculation_settings (id,params_json,created_by,created_at,active) VALUES (?,?,?,?,1)').run(settingsId, JSON.stringify(next), actorId, now()); this.db.exec('COMMIT'); } catch (caught) { this.db.exec('ROLLBACK'); throw caught; }
    this.audit(actorId, 'SETTINGS_UPDATED', 'CALCULATION_SETTINGS', settingsId, next); return { id: settingsId, params: next };
  }
  calculateVersion(versionId, settings = this.activeSettings()) {
    const version = this.db.prepare('SELECT * FROM data_versions WHERE id=?').get(versionId); if (!version) throw new Error('Версия данных не найдена.');
    const products = this.db.prepare('SELECT * FROM version_products WHERE version_id=?').all(versionId); const periods = this.db.prepare('SELECT * FROM sales_periods WHERE version_id=? ORDER BY period_number').all(versionId); const periodDays = version.period_granularity === 'day' ? 1 : version.period_granularity === 'week' ? 7 : 30.4375;
    const rows = products.map((product) => {
      const rawPeriods = periods.filter((item) => item.product_id === product.id); const raw = rawPeriods.map((item) => item.net_qty); const demand = raw.map((value) => Math.max(0, value)); const sorted = [...demand].sort((a, b) => a - b); const q1 = percentile(sorted, .25); const q3 = percentile(sorted, .75); const limit = q3 + 3 * (q3 - q1); const outlierIndexes = demand.map((value, index) => value > limit && demand.length >= 8 ? index : -1).filter((index) => index >= 0); const model = settings.params.excludeOutliers && outlierIndexes.length ? demand.map((value, index) => outlierIndexes.includes(index) ? median(demand) : value) : demand;
      const mean = average(model); const deviation = stddevPopulation(model, mean); const cv = mean > 0 ? deviation / mean * 100 : null; const xyz = mean === 0 ? null : cv <= settings.params.xyzX ? 'X' : cv <= settings.params.xyzY ? 'Y' : 'Z'; const totalSold = demand.reduce((sum, value) => sum + value, 0);
      const revenue = version.financial_type === 'REVENUE' && product.source_financial_total !== null ? product.source_financial_total : totalSold * (product.avg_price ?? 0); const grossProfit = version.financial_type === 'GROSS_PROFIT' && product.source_financial_total !== null ? product.source_financial_total : revenue - totalSold * (product.avg_cost ?? 0);
      const third = Math.max(1, Math.floor(model.length / 3)); const previous = average(model.slice(Math.max(0, model.length - third * 2), Math.max(0, model.length - third))); const latest = average(model.slice(-third)); const demandChange = previous === 0 ? (latest > 0 ? null : 0) : (latest - previous) / previous * 100;
      const averageDaily = mean / periodDays; const deviationDaily = deviation / Math.sqrt(periodDays); const lead = (product.lead_time_days ?? settings.params.defaultLeadDays) + (product.delay_days ?? settings.params.defaultDelayDays); const review = product.review_period_days ?? settings.params.defaultReviewDays; const protection = lead + review; const safety = settings.params.safetyZ * deviationDaily * Math.sqrt(protection); const free = Math.max(0, (product.stock_qty ?? 0) - (product.reserve_qty ?? 0)); const position = free + Math.max(0, product.in_transit_qty ?? 0); const reorder = averageDaily * lead + safety; const target = averageDaily * protection + safety; const qty = Math.ceil(Math.max(0, target - position)); const days = averageDaily > 0 ? free / averageDaily : null;
      const lastPositive = rawPeriods.filter((item) => item.net_qty > 0).at(-1); const lastSale = lastPositive?.period_end || null; const noSalesDays = lastSale ? Math.max(0, Math.floor((Date.now() - new Date(`${lastSale}T23:59:59Z`).getTime()) / 86400000)) : Infinity; let status = 'NORMAL';
      if (free <= 0 && averageDaily > 0) status = 'NO_STOCK'; else if (free > 0 && noSalesDays >= settings.params.obsoleteDays) status = 'OBSOLETE'; else if (averageDaily > 0 && (position <= reorder || days <= lead)) status = 'ORDER_NOW'; else if (averageDaily > 0 && days <= lead + settings.params.earlyWarningDays) status = 'ORDER_SOON'; else if (averageDaily > 0 && days > settings.params.excessDays) status = 'EXCESS'; else if (free > 0 && noSalesDays >= settings.params.slowDays) status = 'SLOW'; else if (model.length < settings.params.minHistory) status = 'INSUFFICIENT_DATA';
      const warnings = []; if (outlierIndexes.length) warnings.push('OUTLIER'); if (product.stock_qty < 0) warnings.push('NEGATIVE_STOCK'); if (product.avg_price === null || product.avg_cost === null) warnings.push('PRICE_OR_COST_MISSING'); if (mean === 0) warnings.push('NO_DEMAND'); if (model.length < settings.params.minHistory) warnings.push('SHORT_HISTORY');
      return { product, totalSold, revenue, grossProfit, xyz, cv, mean, deviation, demandChange, outlierCount: outlierIndexes.length, free, position, lead, safety, reorder, target, qty, days, lastSale, status, warnings, abcRevenue: 'C', abcProfit: 'C', abcQuantity: 'C' };
    });
    const assignAbc = (field, abcField) => { const positive = rows.filter((row) => row[field] > 0).sort((a, b) => b[field] - a[field] || a.product.product_name.localeCompare(b.product.product_name, 'ru')); const total = positive.reduce((sum, row) => sum + row[field], 0); let cumulative = 0; positive.forEach((row) => { const before = cumulative / total * 100; row[abcField] = before < settings.params.abcA ? 'A' : before < settings.params.abcB ? 'B' : 'C'; cumulative += row[field]; }); };
    assignAbc('revenue', 'abcRevenue'); assignAbc('grossProfit', 'abcProfit'); assignAbc('totalSold', 'abcQuantity');
    this.db.exec('BEGIN IMMEDIATE');
    try { this.db.prepare('DELETE FROM analysis_results WHERE version_id=? AND settings_id=?').run(versionId, settings.id); const insert = this.db.prepare('INSERT INTO analysis_results (id,version_id,settings_id,product_id,total_sold,revenue,gross_profit,abc_revenue,abc_profit,abc_quantity,xyz,cv,average_demand,demand_stddev,demand_change_pct,outlier_count,free_stock,stock_position,effective_lead_days,safety_stock,reorder_point,target_stock,recommended_qty,days_of_stock,last_sale_date,primary_status,warnings_json,calculated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'); rows.forEach((row) => insert.run(id(), versionId, settings.id, row.product.id, row.totalSold, row.revenue, row.grossProfit, row.abcRevenue, row.abcProfit, row.abcQuantity, row.xyz, row.cv, row.mean, row.deviation, row.demandChange, row.outlierCount, row.free, row.position, row.lead, row.safety, row.reorder, row.target, row.qty, row.days, row.lastSale, row.status, JSON.stringify(row.warnings), now())); this.db.exec('COMMIT'); } catch (caught) { this.db.exec('ROLLBACK'); throw caught; }
    return { versionId, settingsId: settings.id, rows: rows.length };
  }
  recalculateActive(actorId) { const version = this.activeVersion(); if (!version) throw new Error('Нет активной версии данных для расчёта.'); const result = this.calculateVersion(version.id); this.audit(actorId, 'ANALYSIS_RECALCULATED', 'DATA_VERSION', version.id, { settingsId: result.settingsId, rows: result.rows }); return result; }
  resultRows(versionId = this.activeVersion()?.id) {
    if (!versionId) return [];
    const settings = this.activeSettings();
    const rows = this.db.prepare(`SELECT r.*, p.product_name,p.characteristic_name,p.characteristic_code,p.external_code,p.category,p.supplier,p.stock_qty,p.stock_value,p.reserve_qty,p.in_transit_qty,p.avg_cost,p.avg_price,p.lead_time_days,p.delay_days,p.review_period_days,p.months_with_sales
      FROM analysis_results r JOIN version_products p ON p.id=r.product_id WHERE r.version_id=? AND r.settings_id=?`).all(versionId, settings.id)
      .map((row) => ({ ...row, warnings: JSON.parse(row.warnings_json || '[]') }));
    const adjustments = new Map(this.db.prepare('SELECT * FROM manual_adjustments WHERE version_id=? AND reverted_at IS NULL').all(versionId).map((item) => [item.product_id, item]));
    return rows.map((row) => ({ ...row, manual_adjustment: adjustments.get(row.product_id) || null, final_qty: adjustments.get(row.product_id)?.adjusted_qty ?? row.recommended_qty }));
  }
  results(versionId = this.activeVersion()?.id) { return this.resultRows(versionId); }
  analyticsList(query = {}) {
    const all = this.resultRows(); const values = (key) => String(query[key] || '').split(',').filter(Boolean);
    const search = String(query.search || '').trim().toLocaleLowerCase('ru'); const statuses = values('status'); const abc = values('abc'); const xyz = values('xyz');
    const filtered = all.filter((row) => {
      const haystack = [row.product_name, row.characteristic_name, row.characteristic_code, row.external_code, row.category, row.supplier].join(' ').toLocaleLowerCase('ru');
      return (!search || haystack.includes(search)) && (!statuses.length || statuses.includes(row.primary_status)) && (!abc.length || abc.includes(row.abc_revenue)) && (!xyz.length || xyz.includes(row.xyz || 'NONE')) && (!query.category || row.category === query.category) && (!query.supplier || row.supplier === query.supplier);
    });
    const sort = String(query.sort || 'priority'); const direction = String(query.direction || 'asc') === 'desc' ? -1 : 1;
    const priority = { NO_STOCK: 0, ORDER_NOW: 1, ORDER_SOON: 2, OBSOLETE: 3, SLOW: 4, EXCESS: 5, INSUFFICIENT_DATA: 6, NORMAL: 7 };
    filtered.sort((a, b) => {
      const av = sort === 'name' ? a.product_name : sort === 'stock' ? a.free_stock : sort === 'sales' ? a.total_sold : sort === 'days' ? (a.days_of_stock ?? Infinity) : sort === 'recommendation' ? a.recommended_qty : (priority[a.primary_status] ?? 99);
      const bv = sort === 'name' ? b.product_name : sort === 'stock' ? b.free_stock : sort === 'sales' ? b.total_sold : sort === 'days' ? (b.days_of_stock ?? Infinity) : sort === 'recommendation' ? b.recommended_qty : (priority[b.primary_status] ?? 99);
      return typeof av === 'string' ? av.localeCompare(bv, 'ru') * direction : (av - bv) * direction;
    });
    const pageSize = Math.max(1, Math.min(200, Math.floor(numberOr(query.pageSize, 50)))); const page = Math.max(1, Math.floor(numberOr(query.page, 1))); const total = filtered.length;
    return { items: filtered.slice((page - 1) * pageSize, page * pageSize), total, page, pageSize, filters: { categories: [...new Set(all.map((row) => row.category).filter(Boolean))].sort(), suppliers: [...new Set(all.map((row) => row.supplier).filter(Boolean))].sort() } };
  }
  analyticsSummary() {
    const rows = this.resultRows(); const status = {}; const matrix = {};
    rows.forEach((row) => { status[row.primary_status] = (status[row.primary_status] || 0) + 1; const key = `${row.abc_revenue}${row.xyz || '—'}`; const cell = matrix[key] ||= { key, count: 0, revenue: 0, grossProfit: 0, stockValue: 0 }; cell.count += 1; cell.revenue += row.revenue; cell.grossProfit += row.gross_profit; cell.stockValue += row.stock_value || 0; });
    const priority = { NO_STOCK: 0, ORDER_NOW: 1, ORDER_SOON: 2, OBSOLETE: 3, SLOW: 4, EXCESS: 5, INSUFFICIENT_DATA: 6, NORMAL: 7 };
    const version = this.activeVersion(); const trend = version ? this.db.prepare('SELECT period_number,period_label,SUM(CASE WHEN net_qty > 0 THEN net_qty ELSE 0 END) AS sold FROM sales_periods WHERE version_id=? GROUP BY period_number,period_label ORDER BY period_number').all(version.id) : [];
    return { version, total: rows.length, status, matrix: Object.values(matrix), trend, orderNow: rows.filter((row) => ['NO_STOCK','ORDER_NOW','ORDER_SOON'].includes(row.primary_status)), totalRecommendedQty: rows.reduce((sum, row) => sum + row.recommended_qty, 0), totalRecommendedCost: rows.reduce((sum, row) => sum + row.recommended_qty * (row.avg_cost || 0), 0), topUrgent: [...rows].sort((a, b) => (priority[a.primary_status] - priority[b.primary_status]) || ((a.days_of_stock ?? Infinity) - (b.days_of_stock ?? Infinity))).slice(0, 8), topObsolete: rows.filter((row) => ['OBSOLETE','SLOW','EXCESS'].includes(row.primary_status)).sort((a, b) => (b.stock_value || 0) - (a.stock_value || 0)).slice(0, 8) };
  }
  productAnalytics(productId) {
    const row = this.resultRows().find((item) => item.product_id === productId); if (!row) return null;
    const periods = this.db.prepare('SELECT period_number,period_label,period_start,period_end,net_qty FROM sales_periods WHERE product_id=? ORDER BY period_number').all(productId);
    return { product: row, periods, formula: { averageDemand: row.average_demand, effectiveLeadDays: row.effective_lead_days, safetyStock: row.safety_stock, reorderPoint: row.reorder_point, targetStock: row.target_stock, freeStock: row.free_stock, stockPosition: row.stock_position, recommendedQty: row.recommended_qty } };
  }
  activeAdjustment(versionId, productId) { return this.db.prepare('SELECT a.*,u.name AS created_by_name FROM manual_adjustments a JOIN users u ON u.id=a.created_by WHERE a.version_id=? AND a.product_id=? AND a.reverted_at IS NULL ORDER BY a.created_at DESC LIMIT 1').get(versionId, productId); }
  adjustments(productId, versionId = this.activeVersion()?.id) { return versionId ? this.db.prepare('SELECT a.*,u.name AS created_by_name FROM manual_adjustments a JOIN users u ON u.id=a.created_by WHERE a.version_id=? AND a.product_id=? ORDER BY a.created_at DESC').all(versionId, productId) : []; }
  adjustQuantity(productId, adjustedQty, reason, actorId) {
    const version = this.activeVersion(); if (!version) throw new Error('Нет активной версии данных.'); const result = this.resultRows(version.id).find((row) => row.product_id === productId); if (!result) throw new Error('Товар не найден.');
    const quantity = Number(adjustedQty); const why = normalizeText(reason); if (!Number.isInteger(quantity) || quantity < 0) throw new Error('Количество должно быть целым неотрицательным числом.'); if (why.length < 5 || why.length > 500) throw new Error('Причина изменения должна содержать от 5 до 500 символов.');
    this.db.exec('BEGIN IMMEDIATE'); try { this.db.prepare('UPDATE manual_adjustments SET reverted_at=?,reverted_by=? WHERE version_id=? AND product_id=? AND reverted_at IS NULL').run(now(), actorId, version.id, productId); const adjustment = { id: id(), versionId: version.id, productId, calculatedQty: result.recommended_qty, adjustedQty: quantity, reason: why, actorId, createdAt: now() }; this.db.prepare('INSERT INTO manual_adjustments (id,version_id,product_id,calculated_qty,adjusted_qty,reason,created_by,created_at) VALUES (?,?,?,?,?,?,?,?)').run(adjustment.id, adjustment.versionId, adjustment.productId, adjustment.calculatedQty, adjustment.adjustedQty, adjustment.reason, adjustment.actorId, adjustment.createdAt); this.db.exec('COMMIT'); this.audit(actorId, 'RECOMMENDATION_ADJUSTED', 'MANUAL_ADJUSTMENT', adjustment.id, adjustment); return adjustment; } catch (error) { this.db.exec('ROLLBACK'); throw error; }
  }
  revertAdjustment(productId, actorId) { const version = this.activeVersion(); const current = version && this.activeAdjustment(version.id, productId); if (!current) throw new Error('Ручная корректировка не найдена.'); this.db.prepare('UPDATE manual_adjustments SET reverted_at=?,reverted_by=? WHERE id=?').run(now(), actorId, current.id); this.audit(actorId, 'RECOMMENDATION_REVERTED', 'MANUAL_ADJUSTMENT', current.id, { productId }); return { ok: true }; }
  orderNumber() { const count = this.db.prepare('SELECT count(*) AS count FROM orders').get().count + 1; return `ZT-${new Date().toISOString().slice(0, 10).replaceAll('-', '')}-${String(count).padStart(4, '0')}`; }
  orderDetail(orderId) { const order = this.db.prepare('SELECT o.*,u.name AS created_by_name FROM orders o JOIN users u ON u.id=o.created_by WHERE o.id=?').get(orderId); if (!order) return null; const items = this.db.prepare('SELECT i.*,p.product_name,p.characteristic_name,p.external_code,p.supplier FROM order_items i JOIN version_products p ON p.id=i.product_id WHERE i.order_id=? ORDER BY p.product_name').all(orderId).map((item) => ({ ...item, snapshot: JSON.parse(item.snapshot_json) })); return { ...order, items }; }
  listOrders() { return this.db.prepare('SELECT o.*,u.name AS created_by_name,(SELECT count(*) FROM order_items i WHERE i.order_id=o.id) AS item_count FROM orders o JOIN users u ON u.id=o.created_by ORDER BY o.created_at DESC LIMIT 100').all(); }
  createOrder({ productIds, title, comment, supplier }, actorId) {
    const version = this.activeVersion(); if (!version) throw new Error('Нет активной версии данных.'); const selected = new Set(Array.isArray(productIds) ? productIds : []); const rows = this.resultRows(version.id).filter((row) => selected.has(row.product_id)).map((row) => ({ ...row, adjustment: this.activeAdjustment(version.id, row.product_id) })); const items = rows.filter((row) => (row.adjustment?.adjusted_qty ?? row.recommended_qty) > 0); if (!items.length) throw new Error('Выберите хотя бы одну позицию с количеством больше нуля.');
    const order = { id: id(), number: this.orderNumber(), title: normalizeText(title) || `Заказ ${this.orderNumber()}`, status: 'DRAFT', supplier: normalizeText(supplier) || null, comment: normalizeText(comment) || null, versionId: version.id, settingsId: this.activeSettings().id, actorId, createdAt: now() }; this.db.exec('BEGIN IMMEDIATE'); try { this.db.prepare('INSERT INTO orders (id,number,title,status,supplier,comment,data_version_id,settings_version_id,created_by,created_at) VALUES (?,?,?,?,?,?,?,?,?,?)').run(order.id, order.number, order.title, order.status, order.supplier, order.comment, order.versionId, order.settingsId, order.actorId, order.createdAt); const insert = this.db.prepare('INSERT INTO order_items (id,order_id,product_id,calculated_qty,adjusted_qty,final_qty,unit_cost,snapshot_json) VALUES (?,?,?,?,?,?,?,?)'); items.forEach((row) => { const adjusted = row.adjustment?.adjusted_qty ?? null; const final = adjusted ?? row.recommended_qty; insert.run(id(), order.id, row.product_id, row.recommended_qty, adjusted, final, row.avg_cost, JSON.stringify(row)); }); this.db.prepare('INSERT INTO order_status_history (id,order_id,from_status,to_status,actor_id,comment,created_at) VALUES (?,?,?,?,?,?,?)').run(id(), order.id, null, 'DRAFT', actorId, null, now()); this.db.exec('COMMIT'); this.audit(actorId, 'ORDER_CREATED', 'ORDER', order.id, { number: order.number, itemCount: items.length }); return this.orderDetail(order.id); } catch (error) { this.db.exec('ROLLBACK'); throw error; }
  }
  updateOrder(orderId, { title, comment, items }, actorId) { const order = this.orderDetail(orderId); if (!order) throw new Error('Заказ не найден.'); if (order.status !== 'DRAFT') throw new Error('Сформированный или архивный заказ нельзя изменять.'); this.db.exec('BEGIN IMMEDIATE'); try { this.db.prepare('UPDATE orders SET title=?,comment=? WHERE id=?').run(normalizeText(title) || order.title, normalizeText(comment) || null, orderId); if (Array.isArray(items)) { const update = this.db.prepare('UPDATE order_items SET final_qty=? WHERE id=? AND order_id=?'); items.forEach((item) => { const qty = Number(item.finalQty); if (!Number.isInteger(qty) || qty < 0) throw new Error('Количество заказа должно быть целым неотрицательным числом.'); update.run(qty, item.id, orderId); }); } this.db.exec('COMMIT'); this.audit(actorId, 'ORDER_UPDATED', 'ORDER', orderId, {}); return this.orderDetail(orderId); } catch (error) { this.db.exec('ROLLBACK'); throw error; } }
  changeOrderStatus(orderId, status, actorId) { const order = this.orderDetail(orderId); if (!order) throw new Error('Заказ не найден.'); if (!['FORMED','ARCHIVED'].includes(status)) throw new Error('Некорректный статус заказа.'); if (order.status !== 'DRAFT' && status === 'FORMED') throw new Error('Заказ уже сформирован или архивирован.'); const timestamp = now(); this.db.prepare('UPDATE orders SET status=?,formed_at=CASE WHEN ?=\'FORMED\' THEN ? ELSE formed_at END,archived_at=CASE WHEN ?=\'ARCHIVED\' THEN ? ELSE archived_at END WHERE id=?').run(status, status, timestamp, status, timestamp, orderId); this.db.prepare('INSERT INTO order_status_history (id,order_id,from_status,to_status,actor_id,comment,created_at) VALUES (?,?,?,?,?,?,?)').run(id(), orderId, order.status, status, actorId, null, timestamp); this.audit(actorId, status === 'FORMED' ? 'ORDER_FORMED' : 'ORDER_ARCHIVED', 'ORDER', orderId, {}); return this.orderDetail(orderId); }
  copyOrder(orderId, actorId) { const old = this.orderDetail(orderId); if (!old || old.status !== 'FORMED') throw new Error('Можно копировать только сформированный заказ.'); const version = this.activeVersion(); const order = { id: id(), number: this.orderNumber(), title: `${old.title} — копия`, status: 'DRAFT', supplier: old.supplier, comment: old.comment, versionId: version?.id || old.data_version_id, settingsId: this.activeSettings().id, actorId, createdAt: now() }; this.db.exec('BEGIN IMMEDIATE'); try { this.db.prepare('INSERT INTO orders (id,number,title,status,supplier,comment,data_version_id,settings_version_id,created_by,created_at) VALUES (?,?,?,?,?,?,?,?,?,?)').run(order.id, order.number, order.title, order.status, order.supplier, order.comment, order.versionId, order.settingsId, actorId, order.createdAt); const insert = this.db.prepare('INSERT INTO order_items (id,order_id,product_id,calculated_qty,adjusted_qty,final_qty,unit_cost,snapshot_json) VALUES (?,?,?,?,?,?,?,?)'); old.items.forEach((item) => insert.run(id(), order.id, item.product_id, item.calculated_qty, item.adjusted_qty, item.final_qty, item.unit_cost, item.snapshot_json)); this.db.prepare('INSERT INTO order_status_history (id,order_id,from_status,to_status,actor_id,comment,created_at) VALUES (?,?,?,?,?,?,?)').run(id(), order.id, null, 'DRAFT', actorId, 'Копия сформированного заказа', now()); this.db.exec('COMMIT'); this.audit(actorId, 'ORDER_COPIED', 'ORDER', order.id, { sourceOrderId: orderId }); return this.orderDetail(order.id); } catch (error) { this.db.exec('ROLLBACK'); throw error; } }
  activeVersion() { return this.db.prepare("SELECT * FROM data_versions WHERE status = 'ACTIVE' ORDER BY activated_at DESC LIMIT 1").get(); }
  listVersions() { return this.db.prepare('SELECT v.*, u.name AS created_by_name FROM data_versions v JOIN users u ON u.id=v.created_by ORDER BY v.created_at DESC LIMIT 10').all(); }
  listProfiles() { return this.db.prepare("SELECT p.*, u.name AS created_by_name FROM import_profiles p JOIN users u ON u.id=p.created_by ORDER BY p.status, p.updated_at DESC").all().map((profile) => ({ ...profile, mapping: JSON.parse(profile.mapping_json), periodColumns: JSON.parse(profile.period_columns_json) })); }
  profilesForSignature(signature) { return this.listProfiles().filter((profile) => profile.status === 'ACTIVE' && profile.columns_signature === signature); }
  profileById(profileId) { const profile = this.db.prepare('SELECT * FROM import_profiles WHERE id=?').get(profileId); return profile ? { ...profile, mapping: JSON.parse(profile.mapping_json), periodColumns: JSON.parse(profile.period_columns_json) } : null; }
  saveProfile({ id: profileId, name, sourceType, signature, mapping, periodColumns, granularity, order, financialType, actorId }) {
    const timestamp = now();
    if (profileId) {
      const existing = this.profileById(profileId); if (!existing || existing.status !== 'ACTIVE') throw new Error('Профиль импорта не найден или архивирован.');
      this.db.prepare('UPDATE import_profiles SET name=?,source_type=?,columns_signature=?,mapping_json=?,period_columns_json=?,period_granularity=?,periods_order=?,financial_type=?,version_no=version_no+1,updated_at=? WHERE id=?').run(name, sourceType, signature, JSON.stringify(mapping), JSON.stringify(periodColumns), granularity, order, financialType, timestamp, profileId);
      this.audit(actorId, 'IMPORT_PROFILE_UPDATED', 'IMPORT_PROFILE', profileId, { name }); return profileId;
    }
    if (!normalizeText(name)) throw new Error('Укажите название профиля импорта.');
    const newId = id(); this.db.prepare('INSERT INTO import_profiles (id,name,status,source_type,columns_signature,mapping_json,period_columns_json,period_granularity,periods_order,financial_type,created_by,created_at,updated_at) VALUES (?,?,\'ACTIVE\',?,?,?,?,?,?,?,?,?,?)').run(newId, normalizeText(name), sourceType, signature, JSON.stringify(mapping), JSON.stringify(periodColumns), granularity, order, financialType, actorId, timestamp, timestamp);
    this.audit(actorId, 'IMPORT_PROFILE_CREATED', 'IMPORT_PROFILE', newId, { name: normalizeText(name) }); return newId;
  }
  archiveProfile(profileId, actorId) { const profile = this.profileById(profileId); if (!profile || profile.status !== 'ACTIVE') throw new Error('Профиль импорта не найден или уже архивирован.'); this.db.prepare("UPDATE import_profiles SET status='ARCHIVED', archived_at=?, updated_at=? WHERE id=?").run(now(), now(), profileId); this.audit(actorId, 'IMPORT_PROFILE_ARCHIVED', 'IMPORT_PROFILE', profileId, {}); }
  commitImport({ sourceType, sourceName, sourceHash, rows, issues, mapping, periodColumns, period, financialType, actorId, profileId }) {
    const versionId = id(); const timestamp = now(); const previous = this.activeVersion();
    this.db.exec('BEGIN IMMEDIATE');
    try {
      this.db.prepare("UPDATE data_versions SET status='ARCHIVED' WHERE status='ACTIVE'").run();
      this.db.prepare('INSERT INTO data_versions (id,status,source_type,source_name,source_hash,created_by,created_at,activated_at,parent_version_id,period_granularity,last_period_end,periods_order,row_count,product_count,profile_id,financial_type) VALUES (?,\'ACTIVE\',?,?,?,?,?,?,?,?,?,?,?,?,?,?)').run(versionId, sourceType, sourceName, sourceHash, actorId, timestamp, timestamp, previous?.id || null, period.granularity, period.lastPeriodEnd, period.order, rows.length, rows.length, profileId || null, financialType || 'NONE');
      const insertProduct = this.db.prepare('INSERT INTO version_products (id,version_id,external_code,product_name,normalized_name,characteristic_code,characteristic_name,normalized_characteristic_name,composite_key,category,supplier,stock_qty,stock_value,reserve_qty,in_transit_qty,avg_cost,avg_price,lead_time_days,delay_days,review_period_days,source_financial_total,source_total_sold,months_with_sales) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)');
      const insertPeriod = this.db.prepare('INSERT INTO sales_periods (id,version_id,product_id,period_number,period_label,period_start,period_end,net_qty) VALUES (?,?,?,?,?,?,?,?)');
      for (const row of rows) {
        const productId = id(); insertProduct.run(productId, versionId, row.externalCode, row.productName, row.normalizedName, row.characteristicCode, row.characteristicName, row.normalizedCharacteristicName, row.compositeKey, row.category, row.supplier, row.stockQty, row.stockValue, row.reserveQty, row.inTransitQty, row.avgCost, row.avgPrice, row.leadTimeDays, row.delayDays, row.reviewPeriodDays, row.sourceFinancialTotal, row.sourceTotalSold, row.monthsWithSales);
        row.periods.forEach((item, index) => insertPeriod.run(id(), versionId, productId, index + 1, item.label, item.start, item.end, item.value));
      }
      const insertIssue = this.db.prepare('INSERT INTO imports_issues (id,version_id,row_no,column_name,severity,code,message) VALUES (?,?,?,?,?,?,?)');
      issues.forEach((issue) => insertIssue.run(id(), versionId, issue.rowNo || null, issue.column || null, issue.severity, issue.code, issue.message));
      this.db.exec('COMMIT'); this.audit(actorId, 'IMPORT_ACTIVATED', 'DATA_VERSION', versionId, { sourceType, sourceName, rows: rows.length }); return versionId;
    } catch (caught) { this.db.exec('ROLLBACK'); throw caught; }
  }
  restoreVersion(sourceVersionId, actorId) {
    const source = this.db.prepare('SELECT * FROM data_versions WHERE id=?').get(sourceVersionId); if (!source) throw new Error('Версия данных не найдена.');
    const products = this.db.prepare('SELECT * FROM version_products WHERE version_id=?').all(sourceVersionId); const periods = this.db.prepare('SELECT * FROM sales_periods WHERE version_id=? ORDER BY period_number').all(sourceVersionId); const issues = this.db.prepare('SELECT * FROM imports_issues WHERE version_id=?').all(sourceVersionId);
    const rows = products.map((product) => ({ externalCode: product.external_code, productName: product.product_name, normalizedName: product.normalized_name, characteristicCode: product.characteristic_code, characteristicName: product.characteristic_name, normalizedCharacteristicName: product.normalized_characteristic_name, compositeKey: product.composite_key, category: product.category, supplier: product.supplier, stockQty: product.stock_qty, stockValue: product.stock_value, reserveQty: product.reserve_qty, inTransitQty: product.in_transit_qty, avgCost: product.avg_cost, avgPrice: product.avg_price, leadTimeDays: product.lead_time_days, delayDays: product.delay_days, reviewPeriodDays: product.review_period_days, sourceFinancialTotal: product.source_financial_total, sourceTotalSold: product.source_total_sold, monthsWithSales: product.months_with_sales, periods: periods.filter((period) => period.product_id === product.id).map((period) => ({ label: period.period_label, start: period.period_start, end: period.period_end, value: period.net_qty })) }));
    const copiedIssues = issues.map((issue) => ({ rowNo: issue.row_no, column: issue.column_name, severity: issue.severity, code: issue.code, message: issue.message }));
    const versionId = this.commitImport({ sourceType: 'RESTORE', sourceName: `Восстановление: ${source.source_name}`, sourceHash: source.source_hash, rows, issues: copiedIssues, mapping: {}, periodColumns: [], period: { granularity: source.period_granularity, lastPeriodEnd: source.last_period_end, order: source.periods_order }, financialType: source.financial_type, actorId, profileId: source.profile_id });
    this.audit(actorId, 'IMPORT_RESTORED', 'DATA_VERSION', versionId, { restoredFrom: sourceVersionId }); return versionId;
  }
}

function cookies(request) { return Object.fromEntries(String(request.headers.cookie || '').split(';').filter(Boolean).map((entry) => { const [key, ...value] = entry.trim().split('='); return [key, decodeURIComponent(value.join('='))]; })); }
async function body(request) { const chunks = []; let size = 0; for await (const chunk of request) { size += chunk.length; if (size > 70 * 1024 * 1024) throw new Error('Запрос превышает допустимый размер 70 МБ.'); chunks.push(chunk); } if (!chunks.length) return {}; try { return JSON.parse(Buffer.concat(chunks).toString('utf8')); } catch { throw new Error('Некорректный JSON в запросе.'); } }
function send(response, status, data, headers = {}) { response.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', ...headers }); response.end(JSON.stringify(data)); }
function error(response, status, message, code = 'REQUEST_ERROR') { send(response, status, { code, message }); }
function reportWorkbook(store, reportType) {
  const labels = { ORDER: 'К заказу', SUPPLIERS: 'Заказ по поставщикам', ABC_XYZ: 'ABC XYZ', OBSOLETE: 'Неликвиды', ALL: 'Все товары', FINANCE: 'Финансовая сводка', QUALITY: 'Протокол качества', ORDERS: 'История заказов' }; if (!labels[reportType]) throw new Error('Неизвестный тип отчёта.');
  const version = store.activeVersion(); const rows = store.resultRows(); const columns = ['Номенклатура','Характеристика','Код 1С','Поставщик','ABC выручка','ABC прибыль','ABC количество','XYZ','Статус','Свободный остаток, шт','Дней запаса','Точка заказа','Расчётный заказ, шт','Ручной заказ, шт','Итоговый заказ, шт','Себестоимость, руб','Стоимость заказа, руб'];
  const dataRows = (source) => source.map((row) => [row.product_name,row.characteristic_name || '',row.external_code || '',row.supplier || '',row.abc_revenue,row.abc_profit,row.abc_quantity,row.xyz || '',row.primary_status,row.free_stock,row.days_of_stock,row.reorder_point,row.recommended_qty,row.manual_adjustment?.adjusted_qty ?? '',row.final_qty,row.avg_cost ?? '',(row.final_qty || 0) * (row.avg_cost || 0)]);
  const book = XLSX.utils.book_new(); const add = (name, source) => { const sheet = XLSX.utils.aoa_to_sheet([[labels[reportType], new Date().toLocaleString('ru-RU')], [`Активная версия: ${version?.source_name || 'нет'}; параметры анализа: ${store.activeSettings().id}`], columns, ...dataRows(source)]); sheet['!freeze'] = { ySplit: 3 }; sheet['!autofilter'] = { ref: `A3:Q${Math.max(3, source.length + 3)}` }; sheet['!cols'] = columns.map((header) => ({ wch: Math.max(13, header.length + 2) })); XLSX.utils.book_append_sheet(book, sheet, name.slice(0, 31)); };
  if (reportType === 'ORDER') add('К заказу', rows.filter((row) => row.final_qty > 0)); else if (reportType === 'OBSOLETE') add('Неликвиды', rows.filter((row) => ['OBSOLETE','SLOW','EXCESS'].includes(row.primary_status))); else if (reportType === 'ABC_XYZ') add('ABC XYZ', rows); else if (reportType === 'ALL') add('Все товары', rows); else if (reportType === 'SUPPLIERS') { add('Сводный', rows.filter((row) => row.final_qty > 0)); [...new Set(rows.map((row) => row.supplier || 'Без поставщика'))].forEach((supplier) => add(supplier, rows.filter((row) => (row.supplier || 'Без поставщика') === supplier && row.final_qty > 0))); } else if (reportType === 'FINANCE') { const summary = store.analyticsSummary(); const sheet = XLSX.utils.aoa_to_sheet([[labels.FINANCE, new Date().toLocaleString('ru-RU')], ['Показатель','Значение'], ['Стоимость запасов', rows.reduce((sum,row) => sum + (row.stock_value || 0),0)], ['Стоимость рекомендаций', summary.totalRecommendedCost], ['Рисков неликвида', summary.status.OBSOLETE || 0]]); XLSX.utils.book_append_sheet(book, sheet, 'Финансы'); } else if (reportType === 'QUALITY') { const issues = version ? store.db.prepare('SELECT row_no,column_name,severity,code,message FROM imports_issues WHERE version_id=?').all(version.id) : []; XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Протокол качества', new Date().toLocaleString('ru-RU')], ['Строка','Колонка','Уровень','Код','Сообщение'], ...issues.map((item) => [item.row_no,item.column_name,item.severity,item.code,item.message])]), 'Качество'); } else { const orders = store.listOrders(); XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['История заказов', new Date().toLocaleString('ru-RU')], ['Номер','Название','Статус','Поставщик','Позиций','Автор','Дата'], ...orders.map((item) => [item.number,item.title,item.status,item.supplier,item.item_count,item.created_by_name,item.created_at])]), 'Заказы'); }
  return XLSX.write(book, { type: 'buffer', bookType: 'xlsx' });
}
function smtpConfigured() { return Boolean(process.env.SMTP_HOST && process.env.SMTP_FROM); }
function smtpMessage({ from, recipients, subject, message, attachment, filename }) {
  const boundary = `=_inventory_${id()}`; const encode = (value) => Buffer.from(String(value), 'utf8').toString('base64'); const text = Buffer.from(String(message || ''), 'utf8').toString('base64'); const file = attachment.toString('base64').replace(/.{1,76}/g, '$&\r\n');
  return [`From: ${from}`, `To: ${recipients.join(', ')}`, `Subject: =?UTF-8?B?${encode(subject || 'Отчёт по товарным остаткам')}?=`, 'MIME-Version: 1.0', `Content-Type: multipart/mixed; boundary="${boundary}"`, '', `--${boundary}`, 'Content-Type: text/plain; charset=utf-8', 'Content-Transfer-Encoding: base64', '', text, `--${boundary}`, `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name="${filename}"`, 'Content-Transfer-Encoding: base64', `Content-Disposition: attachment; filename="${filename}"`, '', file, `--${boundary}--`, ''].join('\r\n');
}
async function deliverSmtp({ recipients, subject, message, attachment, filename, socket: existingSocket = null, upgraded = false }) {
  if (!smtpConfigured()) throw new Error('SMTP не настроен: задайте SMTP_HOST и SMTP_FROM.'); const host = process.env.SMTP_HOST; const port = Number(process.env.SMTP_PORT || (process.env.SMTP_SECURE === 'true' ? 465 : 587)); const secure = process.env.SMTP_SECURE === 'true'; const timeout = Number(process.env.SMTP_TIMEOUT_MS || 15000);
  const socket = existingSocket || await new Promise((resolve, reject) => { const candidate = secure ? connectTls({ host, port, servername: host }) : connectNet({ host, port }); candidate.setTimeout(timeout, () => candidate.destroy(new Error('Превышено время ожидания SMTP.'))); candidate.once('connect', () => resolve(candidate)); candidate.once('secureConnect', () => resolve(candidate)); candidate.once('error', reject); });
  let buffer = ''; const responses = []; let wake; socket.on('data', (chunk) => { buffer += chunk.toString('utf8'); while (/^\d{3} [^\r\n]*\r?\n/m.test(buffer)) { const match = buffer.match(/^([0-9]{3})([ -])[^\r\n]*\r?\n/); if (!match) break; const code = Number(match[1]); const lines = []; let complete = false; while (!complete) { const line = buffer.match(/^([0-9]{3})([ -])([^\r\n]*)\r?\n/); if (!line) break; buffer = buffer.slice(line[0].length); lines.push(line[3]); complete = line[2] === ' '; } if (!complete) { buffer = `${match[0]}${buffer}`; break; } responses.push({ code, text: lines.join('\n') }); if (wake) { wake(); wake = null; } } });
  const next = () => new Promise((resolve, reject) => { const timer = setTimeout(() => reject(new Error('SMTP не ответил вовремя.')), timeout); const check = () => { if (responses.length) { clearTimeout(timer); return resolve(responses.shift()); } wake = check; }; check(); }); const command = async (line, allowed = [250]) => { socket.write(`${line}\r\n`); const answer = await next(); if (!allowed.includes(answer.code)) throw new Error(`SMTP ${answer.code}: ${answer.text}`); return answer; };
  try { if (!upgraded) { const greeting = await next(); if (greeting.code !== 220) throw new Error(`SMTP ${greeting.code}: ${greeting.text}`); } const hello = await command(`EHLO ${process.env.SMTP_HELO || 'inventory.local'}`); if (!secure && !upgraded && /STARTTLS/i.test(hello.text)) { await command('STARTTLS', [220]); socket.removeAllListeners('data'); const encrypted = await new Promise((resolve, reject) => { const candidate = connectTls({ socket, servername: host }); candidate.once('secureConnect', () => resolve(candidate)); candidate.once('error', reject); }); return await deliverSmtp({ recipients, subject, message, attachment, filename, socket: encrypted, upgraded: true }); }
    if (process.env.SMTP_USER) await command(`AUTH PLAIN ${Buffer.from(`\u0000${process.env.SMTP_USER}\u0000${process.env.SMTP_PASSWORD || ''}`).toString('base64')}`, [235]); await command(`MAIL FROM:<${process.env.SMTP_FROM}>`); for (const recipient of recipients) await command(`RCPT TO:<${recipient}>`, [250, 251]); await command('DATA', [354]); socket.write(`${smtpMessage({ from: process.env.SMTP_FROM, recipients, subject, message, attachment, filename })}\r\n.\r\n`); const sent = await next(); if (sent.code !== 250) throw new Error(`SMTP ${sent.code}: ${sent.text}`); await command('QUIT', [221]); return { ok: true }; } finally { socket.destroy(); }
}
function sessionCookie(value, secure) { return `sid=${encodeURIComponent(value)}; HttpOnly; SameSite=Lax; Path=/; Max-Age=${8 * 60 * 60}${secure ? '; Secure' : ''}`; }
function clearCookie(secure) { return `sid=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0${secure ? '; Secure' : ''}`; }

export function makeApplication({ databasePath = process.env.DATABASE_PATH || join(here, '.data', 'zapasometr.db'), secureCookies = process.env.NODE_ENV === 'production' } = {}) {
  const store = new Store(databasePath); const loginAttempts = new Map(); const previews = new Map();
  const current = (request) => store.session(cookies(request).sid);
  function requireUser(request, response, { admin = false, business = false, csrf = false } = {}) { const session = current(request); if (!session) { error(response, 401, 'Требуется вход в систему.', 'UNAUTHENTICATED'); return null; } if (admin && session.user.role !== 'ADMIN') { error(response, 403, 'Недостаточно прав.', 'FORBIDDEN'); return null; } if (business && !['BUYER', 'MANAGER', 'FINANCE'].includes(session.user.role)) { error(response, 403, 'Недостаточно прав.', 'FORBIDDEN'); return null; } if (csrf && request.headers['x-csrf-token'] !== session.csrf) { error(response, 403, 'Не пройдена проверка безопасности запроса.', 'CSRF_INVALID'); return null; } return session; }
  async function importSource(input) {
    if (input.type === 'DEMO') return { kind: 'text', content: exampleCsv(), sourceType: 'DEMO', sourceName: 'Демонстрационные данные.csv' };
    if (input.type === 'PASTE') { if (!normalizeText(input.text)) throw new Error('Вставьте таблицу с заголовками.'); return { kind: 'text', content: input.text, sourceType: 'PASTE', sourceName: 'Вставленные данные.csv' }; }
    if (input.type === 'FILE') { if (!input.fileName || !input.contentBase64) throw new Error('Выберите файл для импорта.'); const raw = Buffer.from(input.contentBase64, 'base64'); if (raw.length > 50 * 1024 * 1024) throw new Error('Размер файла превышает 50 МБ.'); const isCsv = /\.csv$/i.test(input.fileName); return { kind: isCsv ? 'text' : 'base64', content: isCsv ? raw.toString('utf8') : input.contentBase64, sourceType: 'FILE', sourceName: input.fileName, sheetName: input.sheetName }; }
    if (input.type === 'GOOGLE') {
      if (!/^https:\/\/docs\.google\.com\/spreadsheets\//.test(input.url || '')) throw new Error('Укажите ссылку на опубликованную Google Таблицу.');
      const parsed = new URL(input.url); const match = parsed.pathname.match(/\/d\/([^/]+)/); if (!match) throw new Error('Не удалось определить идентификатор Google Таблицы.'); const gid = parsed.searchParams.get('gid') || '0'; const csvUrl = `https://docs.google.com/spreadsheets/d/${match[1]}/export?format=csv&gid=${encodeURIComponent(gid)}`;
      const external = await fetch(csvUrl); if (!external.ok) throw new Error('Не удалось скачать Google Таблицу. Проверьте, что она опубликована для чтения.'); const content = await external.text(); if (Buffer.byteLength(content) > 50 * 1024 * 1024) throw new Error('Размер файла превышает 50 МБ.'); return { kind: 'text', content, sourceType: 'GOOGLE', sourceName: 'Google Таблица.csv' };
    }
    throw new Error('Неизвестный источник импорта.');
  }
  function previewResult(preview) { return { previewId: preview.id, sourceType: preview.sourceType, sourceName: preview.sourceName, sheets: preview.parsed.sheets, sheetName: preview.parsed.sheetName, headers: preview.parsed.headers, headerRow: preview.parsed.headerRow, rowCount: preview.parsed.rows.length, sampleRows: preview.parsed.rows.slice(0, 20), suggested: preview.suggested, matchedProfiles: preview.matchedProfiles.map((profile) => ({ id: profile.id, name: profile.name, mapping: profile.mapping, periodColumns: profile.periodColumns, periodGranularity: profile.period_granularity, periodsOrder: profile.periods_order, financialType: profile.financial_type, updatedAt: profile.updated_at })) }; }
  async function serveStatic(request, response, urlPath) { const requested = urlPath === '/' ? '/index.html' : urlPath; const safePath = normalize(requested).replace(/^([.][.][/\\])+/, ''); const file = join(here, 'public', safePath); if (!file.startsWith(join(here, 'public'))) return error(response, 403, 'Доступ запрещён.'); try { const contents = await readFile(file); response.writeHead(200, { 'Content-Type': mimeTypes[extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-store' }); response.end(contents); } catch { response.writeHead(404); response.end('Not found'); } }
  const server = createHttpServer(async (request, response) => {
    const url = new URL(request.url, 'http://localhost'); const path = url.pathname;
    try {
      if (request.method === 'GET' && path === '/api/health') return send(response, 200, { ok: true });
      if (request.method === 'POST' && path === '/api/auth/login') { const key = request.socket.remoteAddress || 'unknown'; const attempt = loginAttempts.get(key) || { count: 0, until: 0 }; if (attempt.until > Date.now()) return error(response, 429, 'Слишком много попыток. Повторите через 15 минут.', 'LOGIN_BLOCKED'); const input = await body(request); const user = store.userByEmail(input.email); if (!user || !user.active || !(await passwordMatches(input.password || '', user.password_hash))) { attempt.count += 1; if (attempt.count >= 5) attempt.until = Date.now() + 15 * 60 * 1000; loginAttempts.set(key, attempt); return error(response, 401, 'Неверный email или пароль.', 'INVALID_CREDENTIALS'); } loginAttempts.delete(key); store.setLastLogin(user.id); store.audit(user.id, 'LOGIN', 'USER', user.id, {}); const session = store.createSession(user.id); return send(response, 200, { user: safeUser(store.userById(user.id)), csrfToken: session.csrf }, { 'Set-Cookie': sessionCookie(session.id, secureCookies) }); }
      if (request.method === 'POST' && path === '/api/auth/logout') { const session = requireUser(request, response, { csrf: true }); if (!session) return; store.deleteSession(cookies(request).sid); store.audit(session.user.id, 'LOGOUT', 'USER', session.user.id, {}); return send(response, 200, { ok: true }, { 'Set-Cookie': clearCookie(secureCookies) }); }
      if (request.method === 'GET' && path === '/api/auth/me') { const session = requireUser(request, response); if (!session) return; return send(response, 200, { user: session.user, csrfToken: session.csrf }); }
      if (request.method === 'POST' && path === '/api/auth/reset-password') { const input = await body(request); await store.resetPassword(input.token, input.password); return send(response, 200, { ok: true }); }
      if (request.method === 'POST' && path === '/api/imports/preview') {
        const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return;
        const input = await body(request); const source = await importSource(input.source || {}); const parsed = parseWorkbook({ ...source, sheetName: input.source?.sheetName });
        if (parsed.rows.length > 50000) throw new Error('Количество строк превышает 50 000.');
        const signature = digest(parsed.headers.map(normalizeKey).join('|')); const suggested = defaultMapping(parsed.headers); const matchedProfiles = store.profilesForSignature(signature);
        if (matchedProfiles.length === 1) { suggested.mapping = matchedProfiles[0].mapping; suggested.periodColumns = matchedProfiles[0].periodColumns; }
        const preview = { id: id(), userId: session.user.id, expiresAt: Date.now() + 30 * 60 * 1000, sourceType: source.sourceType, sourceName: source.sourceName, sourceHash: digest(source.content), parsed, signature, suggested, matchedProfiles };
        previews.set(preview.id, preview); store.audit(session.user.id, 'IMPORT_PREVIEWED', 'IMPORT_PREVIEW', preview.id, { sourceType: source.sourceType, sourceName: source.sourceName }); return send(response, 200, previewResult(preview));
      }
      if (request.method === 'POST' && path === '/api/imports/commit') {
        const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return;
        const input = await body(request); const preview = previews.get(input.previewId); if (!preview || preview.userId !== session.user.id || preview.expiresAt < Date.now()) throw new Error('Предварительный просмотр истёк. Загрузите источник ещё раз.');
        const mapping = input.mapping || preview.suggested.mapping; const periodColumns = input.periodColumns || preview.suggested.periodColumns; const period = input.period || {}; const financialType = input.financialType || 'NONE'; const normalized = normalizeImport(preview.parsed, { mapping, periodColumns, period, financialType });
        if (normalized.rows.length > 50000) normalized.issues.push({ severity: 'ERROR', code: 'ROWS_LIMIT', message: 'Количество строк превышает 50 000.' });
        if (normalized.hasErrors || normalized.issues.some((issue) => issue.severity === 'ERROR')) return send(response, 422, { code: 'IMPORT_VALIDATION_FAILED', message: 'Импорт содержит критические ошибки.', issues: normalized.issues, rowCount: normalized.rows.length });
        let profileId = input.profileId || null;
        if (input.profileAction === 'CREATE') profileId = store.saveProfile({ id: null, name: input.profileName, sourceType: preview.sourceType, signature: preview.signature, mapping, periodColumns, granularity: period.granularity, order: period.order, financialType, actorId: session.user.id });
        if (input.profileAction === 'UPDATE') profileId = store.saveProfile({ id: profileId, name: input.profileName, sourceType: preview.sourceType, signature: preview.signature, mapping, periodColumns, granularity: period.granularity, order: period.order, financialType, actorId: session.user.id });
        const versionId = store.commitImport({ sourceType: preview.sourceType, sourceName: preview.sourceName, sourceHash: preview.sourceHash, rows: normalized.rows, issues: normalized.issues, mapping, periodColumns, period, financialType, actorId: session.user.id, profileId }); const analysis = store.recalculateActive(session.user.id); previews.delete(preview.id);
        return send(response, 201, { versionId, rowCount: normalized.rows.length, warnings: normalized.issues.filter((issue) => issue.severity === 'WARNING'), analysis, activeVersion: store.activeVersion() });
      }
      if (request.method === 'GET' && path === '/api/versions') { const session = requireUser(request, response, { business: true }); if (!session) return; return send(response, 200, { active: store.activeVersion(), versions: store.listVersions() }); }
      const restoreMatch = path.match(/^\/api\/versions\/([^/]+)\/restore$/); if (restoreMatch && request.method === 'POST') { const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return; const versionId = store.restoreVersion(restoreMatch[1], session.user.id); const analysis = store.recalculateActive(session.user.id); return send(response, 201, { versionId, analysis, activeVersion: store.activeVersion() }); }
      if (request.method === 'GET' && path === '/api/import-profiles') { const session = requireUser(request, response, { business: true }); if (!session) return; return send(response, 200, { profiles: store.listProfiles() }); }
      if (request.method === 'POST' && path === '/api/import-profiles') { const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return; const input = await body(request); const profileId = store.saveProfile({ id: input.id || null, name: input.name, sourceType: input.sourceType || 'FILE', signature: input.signature, mapping: input.mapping || {}, periodColumns: input.periodColumns || [], granularity: input.granularity || null, order: input.order || null, financialType: input.financialType || 'NONE', actorId: session.user.id }); return send(response, 201, { profileId }); }
      const archiveProfile = path.match(/^\/api\/import-profiles\/([^/]+)\/archive$/); if (archiveProfile && request.method === 'POST') { const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return; store.archiveProfile(archiveProfile[1], session.user.id); return send(response, 200, { ok: true }); }
      if (request.method === 'GET' && path === '/api/settings') { const session = requireUser(request, response, { business: true }); if (!session) return; return send(response, 200, { active: store.activeSettings(), history: store.settingsHistory() }); }
      if (request.method === 'POST' && path === '/api/settings') { const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return; const input = await body(request); const settings = store.updateSettings(input.params || {}, session.user.id); const version = store.activeVersion(); const analysis = version ? store.recalculateActive(session.user.id) : null; return send(response, 201, { settings, analysis }); }
      if (request.method === 'POST' && path === '/api/analytics/recalculate') { const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return; return send(response, 200, { analysis: store.recalculateActive(session.user.id) }); }
      if (request.method === 'GET' && path === '/api/analytics/summary') { const session = requireUser(request, response, { business: true }); if (!session) return; return send(response, 200, { settings: store.activeSettings(), ...store.analyticsSummary() }); }
      const productAnalytics = path.match(/^\/api\/analytics\/products\/([^/]+)$/); if (productAnalytics && request.method === 'GET') { const session = requireUser(request, response, { business: true }); if (!session) return; const item = store.productAnalytics(productAnalytics[1]); if (!item) return error(response, 404, 'Товар не найден.', 'NOT_FOUND'); return send(response, 200, item); }
      if (request.method === 'GET' && path === '/api/analytics/results') { const session = requireUser(request, response, { business: true }); if (!session) return; const result = store.analyticsList(Object.fromEntries(url.searchParams)); return send(response, 200, { settings: store.activeSettings(), version: store.activeVersion(), ...result, results: result.items }); }
      const adjustmentMatch = path.match(/^\/api\/analytics\/products\/([^/]+)\/adjustment$/); if (adjustmentMatch && request.method === 'GET') { const session = requireUser(request, response, { business: true }); if (!session) return; return send(response, 200, { adjustments: store.adjustments(adjustmentMatch[1]) }); }
      if (adjustmentMatch && request.method === 'POST') { const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return; const input = await body(request); return send(response, 201, { adjustment: store.adjustQuantity(adjustmentMatch[1], input.quantity, input.reason, session.user.id) }); }
      const revertAdjustment = path.match(/^\/api\/analytics\/products\/([^/]+)\/adjustment\/revert$/); if (revertAdjustment && request.method === 'POST') { const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return; return send(response, 200, store.revertAdjustment(revertAdjustment[1], session.user.id)); }
      if (path === '/api/orders' && request.method === 'GET') { const session = requireUser(request, response, { business: true }); if (!session) return; return send(response, 200, { orders: store.listOrders() }); }
      if (path === '/api/orders' && request.method === 'POST') { const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return; return send(response, 201, { order: store.createOrder(await body(request), session.user.id) }); }
      const orderMatch = path.match(/^\/api\/orders\/([^/]+)$/); if (orderMatch && request.method === 'GET') { const session = requireUser(request, response, { business: true }); if (!session) return; const order = store.orderDetail(orderMatch[1]); if (!order) return error(response, 404, 'Заказ не найден.', 'NOT_FOUND'); return send(response, 200, { order }); }
      if (orderMatch && request.method === 'PATCH') { const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return; return send(response, 200, { order: store.updateOrder(orderMatch[1], await body(request), session.user.id) }); }
      const orderStatus = path.match(/^\/api\/orders\/([^/]+)\/(form|archive|copy)$/); if (orderStatus && request.method === 'POST') { const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return; const order = orderStatus[2] === 'copy' ? store.copyOrder(orderStatus[1], session.user.id) : store.changeOrderStatus(orderStatus[1], orderStatus[2] === 'form' ? 'FORMED' : 'ARCHIVED', session.user.id); return send(response, 200, { order }); }
      const reportMatch = path.match(/^\/api\/reports\/([A-Z_]+)\.xlsx$/); if (reportMatch && request.method === 'GET') { const session = requireUser(request, response, { business: true }); if (!session) return; const file = reportWorkbook(store, reportMatch[1]); store.audit(session.user.id, 'REPORT_EXPORTED', 'REPORT', reportMatch[1], { reportType: reportMatch[1] }); response.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': `attachment; filename="${reportMatch[1]}-${new Date().toISOString().slice(0,10)}.xlsx"` }); return response.end(file); }
      if (path === '/api/email-deliveries' && request.method === 'POST') { const session = requireUser(request, response, { business: true, csrf: true }); if (!session) return; const input = await body(request); const recipients = Array.isArray(input.recipients) ? input.recipients.map((item) => String(item).trim()).filter(Boolean) : []; if (!recipients.length || recipients.length > 20 || recipients.some((item) => !/^\S+@\S+\.\S+$/.test(item))) throw new Error('Укажите от 1 до 20 корректных email-адресов.'); if (!['ORDER','SUPPLIERS','ABC_XYZ','OBSOLETE','ALL','FINANCE','QUALITY','ORDERS'].includes(input.reportType)) throw new Error('Выберите тип отчёта.'); const delivery = { id: id(), reportType: input.reportType, recipients, status: 'PENDING', error: null, actorId: session.user.id, createdAt: now() }; store.db.prepare('INSERT INTO email_deliveries (id,report_type,recipients_json,subject,message,status,error,created_by,created_at) VALUES (?,?,?,?,?,?,?,?,?)').run(delivery.id, delivery.reportType, JSON.stringify(recipients), normalizeText(input.subject), normalizeText(input.message), delivery.status, null, delivery.actorId, delivery.createdAt); const attachment = reportWorkbook(store, input.reportType); const filename = `${input.reportType}-${new Date().toISOString().slice(0, 10)}.xlsx`; let lastError; for (let attempt = 1; attempt <= 3; attempt += 1) { try { await deliverSmtp({ recipients, subject: input.subject, message: input.message, attachment, filename }); store.db.prepare('UPDATE email_deliveries SET status=?,sent_at=?,error=NULL WHERE id=?').run('SENT', now(), delivery.id); store.audit(session.user.id, 'REPORT_EMAIL_SENT', 'EMAIL_DELIVERY', delivery.id, { reportType: delivery.reportType, recipients, attempt }); return send(response, 202, { delivery: { ...delivery, status: 'SENT' } }); } catch (caught) { lastError = caught instanceof Error ? caught.message : 'Неизвестная SMTP ошибка.'; if (!/^SMTP 4/.test(lastError)) break; } } store.db.prepare('UPDATE email_deliveries SET status=?,error=? WHERE id=?').run('FAILED', lastError, delivery.id); store.audit(session.user.id, 'REPORT_EMAIL_FAILED', 'EMAIL_DELIVERY', delivery.id, { reportType: delivery.reportType, recipients, error: lastError }); return send(response, 202, { delivery: { ...delivery, status: 'FAILED', error: lastError } }); }
      if (path === '/api/business-audit' && request.method === 'GET') { const session = requireUser(request, response, { business: true }); if (!session) return; if (session.user.role === 'BUYER') return error(response, 403, 'Недостаточно прав.', 'FORBIDDEN'); return send(response, 200, { events: store.db.prepare("SELECT a.*,u.name AS actor_name FROM audit_events a LEFT JOIN users u ON u.id=a.actor_id WHERE a.action IN ('RECOMMENDATION_ADJUSTED','RECOMMENDATION_REVERTED','ORDER_CREATED','ORDER_UPDATED','ORDER_FORMED','ORDER_ARCHIVED','ORDER_COPIED','REPORT_EXPORTED','REPORT_EMAIL_SENT','REPORT_EMAIL_FAILED') ORDER BY a.created_at DESC LIMIT 300").all() }); }
      if (path === '/api/users' && request.method === 'GET') { const session = requireUser(request, response, { admin: true }); if (!session) return; return send(response, 200, { users: store.listUsers() }); }
      if (path === '/api/users' && request.method === 'POST') { const session = requireUser(request, response, { admin: true, csrf: true }); if (!session) return; return send(response, 201, { user: await store.createUser(await body(request), session.user.id) }); }
      const userMatch = path.match(/^\/api\/users\/([^/]+)$/); if (userMatch && request.method === 'PATCH') { const session = requireUser(request, response, { admin: true, csrf: true }); if (!session) return; return send(response, 200, { user: store.updateUser(userMatch[1], await body(request), session.user.id) }); }
      const activeMatch = path.match(/^\/api\/users\/([^/]+)\/(block|unblock)$/); if (activeMatch && request.method === 'POST') { const session = requireUser(request, response, { admin: true, csrf: true }); if (!session) return; return send(response, 200, { user: store.setUserActive(activeMatch[1], activeMatch[2] === 'unblock', session.user.id) }); }
      const resetMatch = path.match(/^\/api\/users\/([^/]+)\/password-reset$/); if (resetMatch && request.method === 'POST') { const session = requireUser(request, response, { admin: true, csrf: true }); if (!session) return; const reset = store.createReset(resetMatch[1], session.user.id); return send(response, 201, { resetUrl: `${url.origin}/?reset=${encodeURIComponent(reset.token)}`, expiresAt: reset.expiresAt }); }
      if (path === '/api/audit' && request.method === 'GET') { const session = requireUser(request, response, { admin: true }); if (!session) return; return send(response, 200, { events: store.listAudit() }); }
      if (path.startsWith('/api/')) return error(response, 404, 'Маршрут не найден.', 'NOT_FOUND'); return serveStatic(request, response, path);
    } catch (caught) { const message = caught instanceof Error ? caught.message : 'Неизвестная ошибка.'; const status = /недействительна|истекла|корректный|не менее|не найден|Укажите|последнего|уже существует|Количество|Размер|строк|файл|лист|таблиц|Границы|Проверьте|Нет активной|Причина|Выберите|Сформированный|архивный|заказ|статус|SMTP/.test(message) ? 400 : 500; return error(response, status, status === 500 ? 'Внутренняя ошибка сервера.' : message, status === 500 ? 'INTERNAL_ERROR' : 'VALIDATION_ERROR'); }
  });
  return { server, store };
}
if (process.argv[1] === fileURLToPath(import.meta.url)) { const port = Number(process.env.PORT || 3000); const { server } = makeApplication(); server.listen(port, () => console.log(`Управление товарными остатками доступно: http://localhost:${port}`)); }
