import assert from 'node:assert/strict';
import { mkdtemp, readFile, stat } from 'node:fs/promises';
import { execFile } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';
import { DatabaseSync } from 'node:sqlite';
import { Store } from '../server.mjs';

const exec = promisify(execFile); const directory = await mkdtemp(join(tmpdir(), 'inventory-ops-')); const databasePath = join(directory, 'source.db');
const store = new Store(databasePath); const admin = await store.createUser({ email: 'ops@example.test', name: 'Оператор', password: 'OpsPass!2026', role: 'ADMIN' }); assert.equal(admin.role, 'ADMIN');
const node = process.execPath; const backup = join(directory, 'backup.db');
const result = await exec(node, ['scripts/backup.mjs', backup], { cwd: process.cwd(), env: { ...process.env, DATABASE_PATH: databasePath } }); assert.match(result.stdout, /Резервная копия создана/); assert.ok((await stat(backup)).size > 0);
const copied = new DatabaseSync(backup); assert.equal(copied.prepare('SELECT count(*) AS count FROM users').get().count, 1); copied.close();
await assert.rejects(() => exec(node, ['scripts/restore-backup.mjs', backup], { cwd: process.cwd(), env: { ...process.env, DATABASE_PATH: join(directory, 'target.db') } }), /перезаписывает/);
const restored = join(directory, 'target.db'); await exec(node, ['scripts/restore-backup.mjs', backup], { cwd: process.cwd(), env: { ...process.env, DATABASE_PATH: restored, CONFIRM_RESTORE: 'yes' } }); const target = new DatabaseSync(restored); assert.equal(target.prepare('SELECT count(*) AS count FROM users').get().count, 1); target.close();
const docker = await readFile('Dockerfile', 'utf8'); const compose = await readFile('docker-compose.yml', 'utf8'); const runbook = await readFile('RUNBOOK.md', 'utf8'); const pkg = JSON.parse(await readFile('package.json', 'utf8')); assert.match(docker, /pnpm install --frozen-lockfile --prod/); assert.match(compose, /zapasometr-data/); assert.match(runbook, /Резервная копия/); assert.match(runbook, /health/); assert.match(pkg.scripts.test, /stage6-check/); assert.match(pkg.scripts.check, /server\.mjs/); console.log('Этап 6 эксплуатация: 16 автоматических проверок пройдены.');
