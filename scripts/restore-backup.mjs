import { copyFile, access } from 'node:fs/promises';
import { join } from 'node:path';

const source = process.argv[2];
const target = process.env.DATABASE_PATH || join(process.cwd(), '.data', 'zapasometr.db');
if (!source) throw new Error('Укажите путь к резервной копии: pnpm restore-backup -- /путь/к/копии.db');
await access(source);
if (process.env.CONFIRM_RESTORE !== 'yes') throw new Error('Восстановление перезаписывает рабочую базу. Повторите с CONFIRM_RESTORE=yes и при остановленном приложении.');
await copyFile(source, target);
console.log(`База восстановлена из: ${source}`);
