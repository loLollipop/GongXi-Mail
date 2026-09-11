import { spawn } from 'node:child_process';
import { readdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';

async function collectTests(directory) {
    const entries = await readdir(directory, { withFileTypes: true });
    const files = await Promise.all(entries.map(async (entry) => {
        const path = `${directory}/${entry.name}`;
        if (entry.isDirectory()) return collectTests(path);
        return entry.isFile() && entry.name.endsWith('.test.ts') ? [path] : [];
    }));
    return files.flat();
}

const sourceRoot = fileURLToPath(new URL('../src', import.meta.url)).replaceAll('\\', '/');
const tests = (await collectTests(sourceRoot)).sort();
if (tests.length === 0) {
    console.error('No test files found under server/src');
    process.exit(1);
}

console.log(`Running ${tests.length} test files`);
const child = spawn(process.execPath, ['--import', 'tsx', '--test', ...tests], {
    stdio: 'inherit',
    windowsHide: true,
});

child.once('error', (error) => {
    console.error(error.message);
    process.exit(1);
});
child.once('exit', (code, signal) => {
    if (signal) {
        console.error(`Test process terminated by ${signal}`);
        process.exit(1);
    }
    process.exit(code ?? 1);
});
