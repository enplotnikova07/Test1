import { copyFile, mkdir } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { DatabaseSync } from 'node:sqlite';

const source = process.env.DATABASE_PATH || join(process.cwd(), '.data', 'zapasometr.db');
const target = process.argv[2] || join(process.cwd(), 'backups', `inventory-${new Date().toISOString().replaceAll(':', '-').slice(0, 19)}.db`);
await mkdir(dirname(target), { recursive: true });
const database = new DatabaseSync(source);
database.exec('PRAGMA wal_checkpoint(TRUNCATE)');
database.close();
await copyFile(source, target);
console.log(`Резервная копия создана: ${target}`);
