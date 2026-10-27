import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { access, cp, lstat, mkdir, mkdtemp, readFile, readdir, rename, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { parseDocument } from 'yaml';
import { connectOneDrive } from './onedrive.mjs';
import { converterRoot, correctedConverter } from './converter-player.mjs';
import { addDownloadControls, downloadLinks, downloadMarker, releaseDownloads } from './downloads.mjs';
import { verifyAndExport } from './presentation-browser.mjs';
import { fontMarker } from './fonts.mjs';

const exec = promisify(execFile);
const slugPattern = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
export const MAX_FILE_BYTES = 100 * 1024 * 1024;
export const MAX_SITE_BYTES = 1_000_000_000;

function nonempty(value) {
    return typeof value === 'string' && value.trim() === value && value.length > 0;
}

export function readConfig(text) {
    const doc = parseDocument(text, { uniqueKeys: true });
    assert.equal(doc.errors.length, 0, 'decks.yml contains invalid YAML or duplicate keys.');
    const config = doc.toJS({ maxAliasCount: 0 });
    assert.ok(config && !Array.isArray(config) && Object.keys(config).length === 1 && Array.isArray(config.decks),
        'decks.yml must contain only a decks list.');
    const names = new Set();
    const ids = new Set();
    return config.decks.map(entry => {
        assert.ok(entry && typeof entry === 'object' && !Array.isArray(entry) &&
            Object.keys(entry).sort().join(',') === 'deck-name,onedrive-id', 'Each deck needs only deck-name and onedrive-id.');
        const name = entry['deck-name'];
        const id = entry['onedrive-id'];
        assert.ok(nonempty(name) && slugPattern.test(name), 'deck-name must be a lowercase, hyphen-separated URL name.');
        assert.ok(nonempty(id), `Missing string onedrive-id for ${name}.`);
        assert.ok(!names.has(name), `Duplicate deck-name: ${name}.`);
        assert.ok(!ids.has(id), `Duplicate onedrive-id for ${name}.`);
        names.add(name);
        ids.add(id);
        return { name, id };
    });
}

export function readState(text) {
    const state = JSON.parse(text);
    assert.ok(Array.isArray(state), '.github/decks-state.json must be a list of previously published decks.');
    const names = new Set();
    for (const deck of state) {
        assert.ok(deck && slugPattern.test(deck.name) && nonempty(deck.name) && nonempty(deck.id) &&
            nonempty(deck.eTag) && !names.has(deck.name), 'Invalid or duplicate previously published deck.');
        names.add(deck.name);
    }
    return state;
}

async function exists(file) {
    try {
        await access(file);
        return true;
    } catch (error) {
        if (error.code !== 'ENOENT') throw error;
        return false;
    }
}

export async function convertDeck(input, output, log = console.log) {
    try {
        const cli = await correctedConverter();
        const conversion = exec(process.execPath, [
            '--input-type=module', '-', input,
            '--output', path.join(output, 'index.html'), '--assets', path.join(output, 'assets'), '--strict'
        ], { timeout: 300_000, maxBuffer: 4 * 1024 * 1024, windowsHide: true });
        conversion.child.stdin.end(cli);
        const { stderr } = await conversion;
        if (stderr.trim()) log(stderr.trim());
    } catch (error) {
        throw new Error(`PPTX2HTML conversion failed: ${error.stderr?.trim() || error.message}`, { cause: error });
    }
}

function escapeHtml(value) {
    return value.replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
}

export function renderIndex(decks) {
    const items = decks.map(deck => {
        const url = `https://segunakinyemi.com/slides/${deck.name}/`;
        return `    <li><a href="${url}">${escapeHtml(deck.title)}</a><br><small>${url}</small><br>${downloadLinks(deck.name)}</li>`;
    }).join('\n');
    return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Slides | Segun Akinyemi</title>
  <style>
    body { max-width: 52rem; margin: 3rem auto; padding: 0 1.25rem; font: 1.1rem/1.6 system-ui, sans-serif; color: #222; background: #fff; }
    a { color: #0092ca; text-underline-offset: .15em; } a:hover { color: #f21368; }
    li { margin: .8rem 0; overflow-wrap: anywhere; } footer { margin-top: 3rem; font-size: .85rem; }
  </style>
</head>
<body>
  <h1>Slides</h1>
  <p><a href="https://segunakinyemi.com/speaking/">Speaking</a></p>
${decks.length ? `  <ul>\n${items}\n  </ul>` : '  <p>No decks have been published yet.</p>'}
  <footer><a href="https://github.com/g21589/PPTX2HTML">PPTX2HTML</a> &middot; <a href="./pptx2html-license.txt">MIT License</a></footer>
</body>
</html>
`;
}

export async function validatePublicFiles(directory, {
    maxFileBytes = MAX_FILE_BYTES, maxSiteBytes = MAX_SITE_BYTES
} = {}) {
    let total = 0;
    async function walk(folder) {
        for (const name of await readdir(folder)) {
            const file = path.join(folder, name);
            const info = await lstat(file);
            assert.ok(!info.isSymbolicLink(), `Public output must not contain symlinks: ${name}.`);
            if (info.isDirectory()) {
                assert.ok(!name.startsWith('.') && name !== 'node_modules', `Unexpected public directory: ${name}.`);
                await walk(file);
            } else {
                assert.ok(info.isFile(), `Unexpected public entry: ${name}.`);
                assert.ok(!/\.(pptx|pptm|ppsx|ppsm|potx|potm|pem|key)$/i.test(name) && !name.startsWith('.'),
                    `Source PowerPoints, credentials, and hidden files must not be published: ${name}.`);
                assert.ok(info.size <= maxFileBytes, `Generated file exceeds GitHub's 100 MiB limit: ${path.relative(directory, file)}.`);
                total += info.size;
            }
        }
    }
    await walk(directory);
    assert.ok(total <= maxSiteBytes, 'Generated site exceeds the GitHub Pages 1 GB limit.');
    return total;
}

async function replacePublication(root, stage) {
    const output = path.join(root, 'decks');
    const backup = path.join(stage, 'previous');
    const hadOutput = await exists(output);
    if (hadOutput) await rename(output, backup);
    try {
        await rename(path.join(stage, 'next'), output);
        await rename(path.join(stage, 'state.json'), path.join(root, '.github', 'decks-state.json'));
    } catch (error) {
        await rm(output, { recursive: true, force: true });
        if (hadOutput) await rename(backup, output);
        throw error;
    }
}

export async function updateDecks({
    root,
    drive,
    force = false,
    convert = convertDeck,
    verify = verifyAndExport,
    downloads = releaseDownloads(),
    license = () => readFile(path.join(converterRoot, 'LICENSE'), 'utf8'),
    tempRoot = tmpdir(),
    log = console.log
}) {
    root = path.resolve(root);
    tempRoot = path.resolve(tempRoot);
    const relativeTemp = path.relative(root, tempRoot);
    assert.ok(relativeTemp === '..' || relativeTemp.startsWith(`..${path.sep}`) || path.isAbsolute(relativeTemp),
        'Temporary PowerPoint downloads must stay outside the repository.');
    const config = readConfig(await readFile(path.join(root, 'decks.yml'), 'utf8'));
    const previous = readState(await readFile(path.join(root, '.github', 'decks-state.json'), 'utf8'));
    const assets = await downloads.list();
    const planned = [];
    for (const deck of config) {
        const item = await drive.getItem(deck.id);
        const old = previous.find(published => published.name === deck.name);
        const htmlFile = path.join(root, 'decks', deck.name, 'index.html');
        const html = await exists(htmlFile) ? await readFile(htmlFile, 'utf8') : '';
        const changed = force || old?.id !== deck.id || old.eTag !== item.eTag ||
            !assets.has(`${deck.name}.pptx`) || !assets.has(`${deck.name}.pdf`) ||
            ![downloadMarker, fontMarker].every(marker => html.includes(marker));
        planned.push({ ...deck, item, changed });
    }
    const removed = previous.filter(deck => !config.some(current => current.name === deck.name));
    const next = planned.map(({ name, id, item }) => ({ name, id, eTag: item.eTag }));
    const index = renderIndex(planned.map(({ name, item }) => ({ name, title: item.name.replace(/\.pptx$/i, '') })));
    const indexFile = path.join(root, 'decks', 'index.html');
    const indexChanged = !await exists(indexFile) || await readFile(indexFile, 'utf8') !== index;
    const changedDecks = planned.filter(deck => deck.changed);
    const changed = changedDecks.length > 0 || removed.length > 0 || JSON.stringify(previous) !== JSON.stringify(next) ||
        indexChanged;
    if (!changed) {
        log('No deck changes. No conversion or publication commit is needed.');
        return { changed: false, updated: [], removed: [] };
    }

    const temporary = await mkdtemp(path.join(tempRoot, 'slides-sources-'));
    let stage;
    try {
        stage = await mkdtemp(path.join(root, '.decks-stage-'));
        const site = path.join(stage, 'next');
        if (await exists(path.join(root, 'decks'))) {
            await cp(path.join(root, 'decks'), site, { recursive: true });
        } else {
            await mkdir(site);
        }
        for (const deck of removed) {
            await rm(path.join(site, deck.name), { recursive: true, force: true });
        }
        const preparedDownloads = [];
        for (const deck of changedDecks) {
            log(`Converting ${deck.name}.`);
            const input = path.join(temporary, `${deck.name}.pptx`);
            const output = path.join(site, deck.name);
            await drive.download(deck.item, input);
            await rm(output, { recursive: true, force: true });
            await convert(input, output, log);
            const html = await readFile(path.join(output, 'index.html'), 'utf8');
            assert.ok(html.includes('id="pp-data"') && html.includes('PPTX2HTML'), `Missing interactive converter output for ${deck.name}.`);
            const pdf = path.join(temporary, `${deck.name}.pdf`);
            await writeFile(path.join(output, 'index.html'), addDownloadControls(html, deck.name));
            await verify(output, pdf, { log });
            preparedDownloads.push({ name: deck.name, pptx: input, pdf });
        }
        await writeFile(path.join(site, 'index.html'), index);
        await writeFile(path.join(site, 'pptx2html-license.txt'), await license());
        const bytes = await validatePublicFiles(site);
        await writeFile(path.join(stage, 'state.json'), JSON.stringify(next, null, 2) + '\n');
        const removedAssets = removed.flatMap(deck => [`${deck.name}.pptx`, `${deck.name}.pdf`]).filter(name => assets.has(name));
        await downloads.publish(preparedDownloads, removedAssets);
        await replacePublication(root, stage);
        log(`Prepared ${changedDecks.length} updated and ${removed.length} removed decks (${bytes} published bytes).`);
        return { changed: true, updated: changedDecks.map(deck => deck.name), removed: removed.map(deck => deck.name) };
    } finally {
        await rm(temporary, { recursive: true, force: true });
        if (stage) await rm(stage, { recursive: true, force: true });
    }
}

async function main() {
    const root = fileURLToPath(new URL('../', import.meta.url));
    // Validate repository inputs before acquiring credentials or making network calls.
    readConfig(await readFile(path.join(root, 'decks.yml'), 'utf8'));
    readState(await readFile(path.join(root, '.github', 'decks-state.json'), 'utf8'));
    const drive = await connectOneDrive();
    await updateDecks({
        root, drive,
        force: process.env.FORCE_UPDATE === 'true',
        tempRoot: process.env.RUNNER_TEMP || tmpdir()
    });
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    main().catch(error => {
        console.error(`Deck update failed: ${error.message}`);
        process.exitCode = 1;
    });
}
