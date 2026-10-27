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
import { addDownloadControls, downloadLinks, downloadMarker, releaseDownloads, releaseAssetUrl,
    mediaAssetPrefix, releasedMediaNames, validateDownloadSize } from './downloads.mjs';
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
        return `    <li class="deck">
      <div class="deck-info">
        <h2><a href="${url}" target="_blank" rel="noopener noreferrer" aria-label="${escapeHtml(deck.title)} (opens in a new tab)">${escapeHtml(deck.title)}<svg class="external-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true" focusable="false"><path d="M15 3h6v6m0-6L10 14M10 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-5"/></svg></a></h2>
        <div class="deck-url"><span class="url-text">${url}</span><button class="copy-url" type="button" data-url="${url}" aria-label="Copy URL for ${escapeHtml(deck.title)}" title="Copy URL" hidden><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true" focusable="false"><rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h3"/></svg></button></div>
        <p class="copy-status" role="status"></p>
      </div>
      <div class="downloads">${downloadLinks(deck.name)}</div>
    </li>`;
    }).join('\n');
    return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Slides | Segun Akinyemi</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font: 1rem/1.6 system-ui, sans-serif; color: #1b2533; background: #f6f8fa; }
    body::before { content: ""; display: block; height: 3px; background: linear-gradient(90deg, #0092ca, #f21368); }
    .page { width: min(100% - 2rem, 68rem); margin: 0 auto; }
    a { color: #006f9b; text-underline-offset: .2em; }
    a:hover { color: #b40e4e; }
    a:focus-visible, button:focus-visible { outline: 3px solid #006f9b; outline-offset: 4px; border-radius: 3px; }
    header { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: .5rem 1rem; padding: 1.25rem 0; border-bottom: 1px solid #dce3e9; }
    header a { display: inline-flex; align-items: center; min-height: 44px; }
    .site-name { color: #1b2533; font-weight: 700; text-decoration: none; }
    main { padding: clamp(2rem, 5vw, 4rem) 0; }
    .heading { display: flex; flex-wrap: wrap; align-items: baseline; gap: .5rem 1rem; margin-bottom: 2rem; }
    h1 { margin: 0; font-size: clamp(2.5rem, 6vw, 3.75rem); line-height: 1.15; letter-spacing: -.04em; }
    .count { margin: 0; color: #596573; font-size: .9rem; }
    .deck-list { display: grid; gap: 1rem; margin: 0; padding: 0; list-style: none; }
    .deck { display: grid; grid-template-columns: minmax(0, 1fr) auto; align-items: center; gap: 1rem 2rem; padding: 1.5rem; border: 1px solid #dce3e9; border-radius: 12px; background: #fff; }
    .deck-info { min-width: 0; }
    h2 { margin: 0; font-size: 1.1rem; line-height: 1.5; font-weight: 650; overflow-wrap: anywhere; }
    h2 a { text-decoration: underline; text-decoration-thickness: 1px; }
    h2 a:hover { text-decoration-thickness: 2px; }
    .external-icon { width: .9em; height: .9em; margin-left: .35em; vertical-align: -.1em; }
    .deck-url { display: flex; align-items: center; gap: .35rem; margin: .5rem 0 0; color: #596573; font-size: .8rem; }
    .url-text { min-width: 0; overflow-wrap: anywhere; user-select: all; }
    .copy-url { flex: none; width: 44px; height: 44px; padding: .5rem; border: 0; border-radius: 6px; color: #006f9b; background: transparent; cursor: pointer; }
    .copy-url:hover { background: #eaf3f8; }
    .copy-url svg { vertical-align: middle; }
    .copy-status { margin: .25rem 0 0; font-size: .8rem; overflow-wrap: anywhere; }
    .copy-status:empty { display: none; }
    .downloads { display: flex; flex-wrap: wrap; gap: .5rem; }
    .downloads a { display: inline-flex; align-items: center; justify-content: center; min-height: 44px; max-width: 100%; padding: .65rem .85rem; border: 1px solid #b8cbd7; border-radius: 7px; font-size: .85rem; font-weight: 600; line-height: 1.35; text-align: center; text-decoration: none; }
    .downloads a:first-child { color: #fff; background: #006f9b; border-color: #006f9b; }
    .downloads a:hover { color: #fff; background: #005675; border-color: #005675; }
    .empty { padding: 2rem; border: 1px dashed #b8cbd7; border-radius: 12px; background: #fff; }
    footer { padding: 1.5rem 0 2rem; border-top: 1px solid #dce3e9; color: #596573; font-size: .85rem; }
    @media (max-width: 60rem) {
      .deck { grid-template-columns: minmax(0, 1fr); padding: 1.25rem; }
    }
  </style>
</head>
<body>
  <div class="page">
  <header><a class="site-name" href="https://segunakinyemi.com/">Segun Akinyemi</a><a href="https://segunakinyemi.com/speaking/">Speaking</a></header>
  <main>
  <div class="heading"><h1>Slides</h1><p class="count">${decks.length} ${decks.length === 1 ? 'deck' : 'decks'}</p></div>
${decks.length ? `  <ul class="deck-list" role="list">\n${items}\n  </ul>` : '  <p class="empty">No decks have been published yet.</p>'}
  </main>
  <footer><a href="https://github.com/g21589/PPTX2HTML">PPTX2HTML</a> &middot; <a href="./pptx2html-license.txt">MIT License</a></footer>
  </div>
  <script>
    document.querySelectorAll('.copy-url').forEach(button => {
      button.hidden = false;
      button.addEventListener('click', async () => {
        const status = button.closest('.deck-info').querySelector('.copy-status');
        status.textContent = '';
        try {
          await navigator.clipboard.writeText(button.dataset.url);
          status.textContent = 'Copied!';
        } catch (error) {
          console.error('Could not copy slide URL.', error);
          status.textContent = 'Copy failed. Select the URL and copy it manually.';
        }
      });
    });
  </script>
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

export async function prepareReleaseMedia(directory, name, temporary, maxFileBytes = MAX_FILE_BYTES) {
    await validatePublicFiles(directory, { maxFileBytes: Infinity, maxSiteBytes: Infinity });
    const folder = path.join(directory, 'assets');
    const media = [];
    if (!await exists(folder)) return media;
    const htmlFile = path.join(directory, 'index.html');
    let html = await readFile(htmlFile, 'utf8');
    // The converter's asset writer emits a flat directory of uniquely named media files.
    for (const entry of await readdir(folder, { withFileTypes: true })) {
        if (!entry.isFile()) continue;
        const source = path.join(folder, entry.name);
        const info = await lstat(source);
        if (info.size <= maxFileBytes) continue;
        validateDownloadSize(info.size, entry.name);
        assert.match(entry.name, /^[a-zA-Z0-9._-]+$/, `Unsupported release media filename: ${entry.name}.`);
        const destination = path.join(temporary, mediaAssetPrefix(name) + entry.name);
        const local = `assets/${encodeURIComponent(entry.name)}`;
        const pattern = new RegExp(local.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?=["\'\\s)\\\\]|$)', 'g');
        const updated = html.replace(pattern, () => releaseAssetUrl(path.basename(destination)));
        assert.notEqual(updated, html, `Could not locate references to oversized media: ${entry.name}.`);
        await cp(source, destination);
        await rm(source);
        html = updated;
        media.push(destination);
    }
    if (media.length) await writeFile(htmlFile, html);
    return media;
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
            releasedMediaNames(html, deck.name).some(name => !assets.has(name)) ||
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
            const media = await prepareReleaseMedia(output, deck.name, temporary);
            preparedDownloads.push({ name: deck.name, pptx: input, pdf, media });
        }
        await writeFile(path.join(site, 'index.html'), index);
        await writeFile(path.join(site, 'pptx2html-license.txt'), await license());
        const bytes = await validatePublicFiles(site);
        await writeFile(path.join(stage, 'state.json'), JSON.stringify(next, null, 2) + '\n');
        const currentMedia = new Set(preparedDownloads.flatMap(file => file.media.map(file => path.basename(file))));
        const removedAssets = [...assets].filter(name =>
            removed.some(deck => name === `${deck.name}.pptx` || name === `${deck.name}.pdf` || name.startsWith(mediaAssetPrefix(deck.name))) ||
            changedDecks.some(deck => name.startsWith(mediaAssetPrefix(deck.name))) && !currentMedia.has(name));
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
