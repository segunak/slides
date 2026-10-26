import { execFile } from 'node:child_process';
import { cp, mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { converterRoot, correctedPlayer, correctTimingSource, correctEngineSource } from './converter-player.mjs';

const exec = promisify(execFile);
const root = fileURLToPath(new URL('../', import.meta.url));
const temporary = await mkdtemp(path.join(tmpdir(), 'slides-converter-check-'));
try {
    const upstream = path.join(temporary, 'pptx2html');
    await cp(converterRoot, upstream, { recursive: true });
    await symlink(path.join(root, 'node_modules'), path.join(upstream, 'node_modules'), process.platform === 'win32' ? 'junction' : 'dir');
    await writeFile(path.join(upstream, 'src', 'generated', 'player-source.js'),
        `export const PLAYER_SOURCE = ${JSON.stringify(await correctedPlayer())};\n`);
    for (const [folder, name, correct] of [['pptx', 'timing', correctTimingSource], ['runtime', 'engine', correctEngineSource]]) {
        const file = path.join(upstream, 'src', folder, `${name}.js`);
        await writeFile(file, correct(await readFile(file, 'utf8')));
    }
    const { stdout, stderr } = await exec(process.execPath, [
        path.join(upstream, 'tools', 'verify.mjs'), '--no-lo', '--out', path.join(temporary, 'report')
    ], { timeout: 300_000, maxBuffer: 8 * 1024 * 1024, windowsHide: true });
    if (stdout.trim()) console.log(stdout.trim());
    if (stderr.trim()) console.error(stderr.trim());
} catch (error) {
    console.error(`Converter verification failed: ${error.stdout?.trim() || error.message}`);
    if (error.stderr?.trim()) console.error(error.stderr.trim());
    process.exitCode = 1;
} finally {
    await rm(temporary, { recursive: true, force: true });
}
