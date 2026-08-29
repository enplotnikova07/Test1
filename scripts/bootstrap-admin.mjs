import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Store } from '../server.mjs';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const [email, password, ...nameParts] = process.argv.slice(2);
if (!email || !password || !nameParts.length) {
  console.error('Использование: node scripts/bootstrap-admin.mjs <email> <пароль> <имя администратора>');
  process.exit(1);
}
const store = new Store(join(root, '.data', 'zapasometr.db'));
if (store.activeAdminCount() > 0) {
  console.error('Активный администратор уже существует. Создайте пользователя через интерфейс.');
  process.exit(1);
}
const admin = await store.createUser({ email, password, name: nameParts.join(' '), role: 'ADMIN' });
console.log(`Создан администратор ${admin.email}.`);
