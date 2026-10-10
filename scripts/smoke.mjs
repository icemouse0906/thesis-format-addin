import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
const result = spawnSync(process.execPath, ['--test', 'tests/regression.test.cjs'], {
  cwd: fileURLToPath(new URL('..', import.meta.url)), stdio: 'inherit',
});
if (result.error) console.error(result.error.message);
process.exit(result.status ?? 1);
