import assert from 'node:assert/strict';
import { crc32 } from 'node:zlib';
import { access, mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { test } from 'node:test';
import { chromium } from 'playwright';
import { convertDeck, readConfig, readState, renderIndex, updateDecks, validatePublicFiles, MAX_FILE_BYTES, MAX_SITE_BYTES } from './update-decks.mjs';
import { addDownloadControls } from './downloads.mjs';
import { verifyAndExport } from './presentation-browser.mjs';

const html = '<!doctype html><html><head><meta name="generator" content="PPTX2HTML"><script id="pp-fonts" type="application/json">[]</script></head><body><script id="pp-data">{}</script></body></html>';
const config = decks => `decks:\n${decks.map(([name, id]) => `  - deck-name: "${name}"\n    onedrive-id: "${id}"`).join('\n')}\n`;
const published = (name, id = name, eTag = 'v1') => ({ name, id, eTag });

async function fixture(t, old = [], selected = old.map(deck => [deck.name, deck.id])) {
    const directory = await mkdtemp(path.join(tmpdir(), 'slides-update-test-'));
    t.after(() => rm(directory, { recursive: true, force: true }));
    const root = path.join(directory, 'repo');
    const temporary = path.join(directory, 'temporary');
    await mkdir(path.join(root, '.github'), { recursive: true });
    await mkdir(path.join(root, 'decks'));
    await mkdir(temporary);
    await writeFile(path.join(root, 'decks.yml'), selected.length ? config(selected) : 'decks: []\n');
    await writeFile(path.join(root, '.github', 'decks-state.json'), JSON.stringify(old, null, 2) + '\n');
    await writeFile(path.join(root, 'decks', 'index.html'), renderIndex(old.map(deck => ({ name: deck.name, title: deck.id }))));
    for (const deck of old) {
        await mkdir(path.join(root, 'decks', deck.name, 'assets'), { recursive: true });
        await writeFile(path.join(root, 'decks', deck.name, 'index.html'), addDownloadControls(`${html}<!-- ${deck.id} -->`, deck.name));
        await writeFile(path.join(root, 'decks', deck.name, 'assets', 'old.png'), 'image');
    }
    const calls = [];
    const assets = new Set(old.flatMap(deck => [`${deck.name}.pptx`, `${deck.name}.pdf`]));
    const options = {
        root, tempRoot: temporary,
        drive: {
            getItem: async id => {
                calls.push(['metadata', id]);
                return { id, name: `${id}.pptx`, eTag: 'v1' };
            },
            download: async (item, destination) => {
                calls.push(['download', item.id, destination]);
                assert.ok(!destination.startsWith(root + path.sep), 'PPTX source must stay outside the checkout.');
                await writeFile(destination, 'temporary pptx');
            }
        },
        convert: async (input, output) => {
            calls.push(['convert', input, output]);
            await mkdir(path.join(output, 'assets'), { recursive: true });
            await writeFile(path.join(output, 'index.html'), html);
            await writeFile(path.join(output, 'assets', 'new.gif'), 'GIF89a');
        },
        verify: async (output, pdf) => {
            calls.push(['verify', output]);
            await writeFile(pdf, '%PDF-fixture');
        },
        downloads: {
            list: async () => assets,
            publish: async (files, removed) => {
                for (const file of files) {
                    assert.equal(await readFile(file.pptx, 'utf8'), 'temporary pptx');
                    assert.equal(await readFile(file.pdf, 'utf8'), '%PDF-fixture');
                }
                for (const file of files) {
                    assets.add(`${file.name}.pptx`);
                    assets.add(`${file.name}.pdf`);
                }
                for (const name of removed) assets.delete(name);
                calls.push(['upload', files.map(file => file.name), removed]);
            }
        },
        license: async () => 'MIT license fixture',
        log: message => calls.push(['log', message])
    };
    return { root, temporary, calls, options, assets };
}

async function snapshot(folder) {
    const entries = {};
    for (const item of await readdir(folder, { withFileTypes: true })) {
        const file = path.join(folder, item.name);
        entries[item.name] = item.isDirectory() ? await snapshot(file) : await readFile(file, 'utf8');
    }
    return entries;
}

test('reads the exact two-string configuration shape and supports an empty list', () => {
    assert.deepEqual(readConfig(config([['alpha', 'id!1']])), [{ name: 'alpha', id: 'id!1' }]);
    assert.deepEqual(readConfig('decks: []'), []);
});

for (const [name, text] of [
    ['mapping instead of list', 'decks: {alpha: id}'],
    ['missing list', 'other: []'],
    ['unknown top-level field', 'decks: []\nother: true'],
    ['duplicate top-level key', 'decks: []\ndecks: []'],
    ['wrong property', 'decks:\n  - name: alpha\n    onedrive-id: id'],
    ['non-string ID', 'decks:\n  - deck-name: alpha\n    onedrive-id: 123'],
    ['blank ID', config([['alpha', ' ']])],
    ['non-string name', 'decks:\n  - deck-name: 123\n    onedrive-id: id'],
    ['duplicate name', config([['alpha', '1'], ['alpha', '2']])],
    ['duplicate ID', config([['alpha', '1'], ['beta', '1']])],
    ['path traversal', config([['../outside', '1']])],
    ['uppercase name', config([['Alpha', '1']])],
    ['extra field', 'decks:\n  - deck-name: alpha\n    onedrive-id: id\n    extra: true'],
    ['alias', 'decks:\n  - &deck {deck-name: alpha, onedrive-id: id}\n  - *deck']
]) {
    test(`rejects configuration with ${name}`, () => assert.throws(() => readConfig(text)));
}

test('reads the simple published-deck list and rejects unsafe directory names', () => {
    assert.deepEqual(readState('[]'), []);
    assert.deepEqual(readState(JSON.stringify([published('alpha')])), [published('alpha')]);
    for (const invalid of [
        null, {}, [published('../outside')], [published('alpha'), published('alpha')], [{ name: 'alpha' }]
    ]) assert.throws(() => readState(JSON.stringify(invalid)));
});

test('unchanged decks are a true publication no-op', async t => {
    const f = await fixture(t, [published('alpha')]);
    const before = await snapshot(f.root);
    assert.deepEqual(await updateDecks(f.options), { changed: false, updated: [], removed: [] });
    assert.deepEqual(await snapshot(f.root), before);
    assert.ok(!f.calls.some(call => call[0] === 'download' || call[0] === 'convert'));
    assert.deepEqual(await readdir(f.temporary), []);
});

test('the HTML index lists every published deck with a visible full URL and clickable title', () => {
    const page = renderIndex([
        { name: 'alpha', title: 'Alpha' },
        { name: 'how-to-train-your-ai', title: 'How to Train Your AI' }
    ]);
    const urls = [...page.matchAll(/<h2><a href="([^"]+)">([^<]+)<\/a><\/h2><p class="deck-url">([^<]+)<\/p>/g)];
    assert.deepEqual(urls.map(match => match.slice(1)), [
        ['https://segunakinyemi.com/slides/alpha/', 'Alpha', 'https://segunakinyemi.com/slides/alpha/'],
        ['https://segunakinyemi.com/slides/how-to-train-your-ai/', 'How to Train Your AI', 'https://segunakinyemi.com/slides/how-to-train-your-ai/']
    ]);
    assert.match(renderIndex([]), /No decks have been published yet/);
    assert.doesNotMatch(renderIndex([]), /<li>/);
});

test('the HTML index is readable and keyboard accessible on desktop, mobile, and enlarged text without scripts', async t => {
    const browser = await chromium.launch();
    t.after(() => browser.close());
    const page = await browser.newPage({ javaScriptEnabled: false });
    const decks = [
        { name: 'context-engineering-is-software-engineering', title: 'Context Engineering Is Software Engineering' },
        { name: 'agentic-ai-technical-audience', title: '[Technical Audiences - Guess What the Agent Did Game] - Agentic AI - From Acronyms to Application' }
    ];
    await page.setContent(renderIndex(decks));
    assert.equal(await page.getByRole('main').count(), 1);
    assert.equal(await page.getByRole('heading', { level: 1 }).textContent(), 'Slides');
    assert.deepEqual(await page.getByRole('heading', { level: 2 }).allTextContents(), decks.map(deck => deck.title));
    assert.equal(await page.getByRole('link', { name: 'Download PDF', exact: true }).count(), decks.length);
    assert.equal(await page.getByRole('link', { name: 'Download PowerPoint', exact: true }).count(), decks.length);
    await page.keyboard.press('Tab');
    assert.equal(await page.locator(':focus').textContent(), 'Segun Akinyemi');
    assert.equal(await page.locator(':focus').evaluate(e => getComputedStyle(e).outlineStyle), 'solid');
    for (const width of [320, 390, 768, 1280]) {
        await page.setViewportSize({ width, height: 900 });
        for (const fontSize of ['100%', '200%']) {
            await page.locator('html').evaluate((e, size) => { e.style.fontSize = size; }, fontSize);
            assert.ok(await page.locator('html').evaluate(e => e.scrollWidth <= innerWidth),
                `Index must not scroll horizontally at ${width}px and ${fontSize} text.`);
            for (const link of await page.locator('.downloads a').all()) {
                const box = await link.boundingBox();
                assert.ok(box && box.height >= 44 && box.x >= 0 && box.x + box.width <= width,
                    'Download buttons must remain usable and inside the viewport.');
            }
        }
    }
    await page.setContent(renderIndex([]));
    assert.equal(await page.getByRole('heading', { level: 1 }).count(), 1);
    assert.equal(await page.getByRole('listitem').count(), 0);
    assert.match(await page.getByRole('main').textContent(), /No decks have been published yet/);
});

for (const condition of ['missing', 'stale']) {
    test(`repairs a ${condition} HTML index without reconverting unchanged decks`, async t => {
        const f = await fixture(t, [published('alpha'), published('beta')]);
        const file = path.join(f.root, 'decks', 'index.html');
        if (condition === 'missing') await rm(file);
        else await writeFile(file, '<a href="./removed-deck/">Stale deck</a>');
        const oldDeck = await snapshot(path.join(f.root, 'decks', 'alpha'));
        const oldState = await readFile(path.join(f.root, '.github', 'decks-state.json'), 'utf8');
        assert.deepEqual(await updateDecks(f.options), { changed: true, updated: [], removed: [] });
        assert.equal(await readFile(file, 'utf8'), renderIndex([
            { name: 'alpha', title: 'alpha' }, { name: 'beta', title: 'beta' }
        ]));
        assert.deepEqual(await snapshot(path.join(f.root, 'decks', 'alpha')), oldDeck);
        assert.equal(await readFile(path.join(f.root, '.github', 'decks-state.json'), 'utf8'), oldState);
        assert.ok(!f.calls.some(call => call[0] === 'download' || call[0] === 'convert'));
        assert.deepEqual(await updateDecks(f.options), { changed: false, updated: [], removed: [] });
    });
}

test('refreshes an outdated title in the index even when the file revision is unchanged', async t => {
    const f = await fixture(t, [published('alpha')]);
    f.options.drive.getItem = async id => ({ id, name: 'Updated Title.pptx', eTag: 'v1' });
    assert.deepEqual(await updateDecks(f.options), { changed: true, updated: [], removed: [] });
    assert.equal(await readFile(path.join(f.root, 'decks', 'index.html'), 'utf8'),
        renderIndex([{ name: 'alpha', title: 'Updated Title' }]));
    assert.ok(!f.calls.some(call => call[0] === 'download' || call[0] === 'convert'));
    assert.deepEqual(await updateDecks(f.options), { changed: false, updated: [], removed: [] });
});

test('converts new and edited decks only, retaining untouched output byte-for-byte', async t => {
    const f = await fixture(t, [published('alpha'), published('beta')], [['alpha', 'alpha'], ['beta', 'beta'], ['gamma', 'gamma']]);
    const oldBeta = await snapshot(path.join(f.root, 'decks', 'beta'));
    f.options.drive.getItem = async id => ({ id, name: `${id}.pptx`, eTag: id === 'alpha' ? 'v2' : 'v1' });
    assert.deepEqual(await updateDecks(f.options), { changed: true, updated: ['alpha', 'gamma'], removed: [] });
    assert.deepEqual(f.calls.filter(call => call[0] === 'download').map(call => call[1]), ['alpha', 'gamma']);
    assert.deepEqual(await snapshot(path.join(f.root, 'decks', 'beta')), oldBeta);
    await assert.rejects(access(path.join(f.root, 'decks', 'alpha', 'assets', 'old.png')));
    assert.deepEqual(readState(await readFile(path.join(f.root, '.github', 'decks-state.json'), 'utf8')), [
        published('alpha', 'alpha', 'v2'), published('beta'), published('gamma')
    ]);
    assert.equal(await readFile(path.join(f.root, 'decks', 'index.html'), 'utf8'), renderIndex([
        { name: 'alpha', title: 'alpha' }, { name: 'beta', title: 'beta' }, { name: 'gamma', title: 'gamma' }
    ]));
    assert.deepEqual(await readdir(f.temporary), []);
    assert.ok(!(await readdir(f.root)).some(name => name.startsWith('.decks-stage-')));
    await validatePublicFiles(path.join(f.root, 'decks'));
    assert.deepEqual(f.calls.find(call => call[0] === 'upload').slice(1), [['alpha', 'gamma'], []]);
    assert.ok(f.calls.findIndex(call => call[0] === 'upload') > f.calls.findLastIndex(call => call[0] === 'verify'));
    const page = await readFile(path.join(f.root, 'decks', 'alpha', 'index.html'), 'utf8');
    assert.match(page, /data-slide-downloads/);
    assert.match(page, /releases\/download\/downloads\/alpha\.pptx/);
});

test('rebuilds a renamed URL or changed source ID, without retaining the old URL', async t => {
    const f = await fixture(t, [published('old-name', 'original')], [['new-name', 'replacement']]);
    assert.deepEqual(await updateDecks(f.options), { changed: true, updated: ['new-name'], removed: ['old-name'] });
    await assert.rejects(access(path.join(f.root, 'decks', 'old-name')));
    const page = await readFile(path.join(f.root, 'decks', 'index.html'), 'utf8');
    assert.equal(page, renderIndex([{ name: 'new-name', title: 'replacement' }]));
    assert.doesNotMatch(page, /old-name/);
    assert.deepEqual(f.calls.find(call => call[0] === 'upload').slice(1), [['new-name'], ['old-name.pptx', 'old-name.pdf']]);
});

test('a replacement source under the same URL rebuilds even if its eTag matches', async t => {
    const f = await fixture(t, [published('alpha', 'original')], [['alpha', 'replacement']]);
    assert.deepEqual(await updateDecks(f.options), { changed: true, updated: ['alpha'], removed: [] });
    assert.deepEqual(readState(await readFile(path.join(f.root, '.github', 'decks-state.json'), 'utf8')), [
        published('alpha', 'replacement')
    ]);
    assert.match(await readFile(path.join(f.root, 'decks', 'index.html'), 'utf8'), />replacement<\/a>/);
});

test('removing entries removes generated directories without downloading any source', async t => {
    const f = await fixture(t, [published('alpha')], []);
    assert.deepEqual(await updateDecks(f.options), { changed: true, updated: [], removed: ['alpha'] });
    assert.ok(!f.calls.some(call => call[0] === 'download'));
    assert.deepEqual(readState(await readFile(path.join(f.root, '.github', 'decks-state.json'), 'utf8')), []);
    assert.equal(await readFile(path.join(f.root, 'decks', 'index.html'), 'utf8'), renderIndex([]));
    await assert.rejects(access(path.join(f.root, 'decks', 'alpha')));
});

for (const change of ['force', 'missing-output']) {
    test(`rebuilds for ${change}`, async t => {
        const f = await fixture(t, [published('alpha')]);
        if (change === 'force') f.options.force = true;
        if (change === 'missing-output') await rm(path.join(f.root, 'decks', 'alpha'), { recursive: true });
        assert.deepEqual((await updateDecks(f.options)).updated, ['alpha']);
    });
}

for (const failure of ['metadata', 'download', 'conversion', 'verification', 'upload', 'invalid-output', 'source-in-output']) {
    test(`${failure} failure leaves all previously published output and state untouched`, async t => {
        const f = await fixture(t, [published('alpha'), published('beta')],
            [['alpha', 'alpha'], ['beta', 'beta'], ['gamma', 'gamma']]);
        f.options.force = true;
        const before = await snapshot(f.root);
        if (failure === 'metadata') f.options.drive.getItem = async () => { throw new Error('Metadata HTTP 404'); };
        if (failure === 'download') f.options.drive.download = async () => { throw new Error('Download failed'); };
        if (failure === 'verification') f.options.verify = async () => { throw new Error('Revealed text missing'); };
        if (failure === 'upload') f.options.downloads.publish = async () => { throw new Error('Release upload failed'); };
        if (failure === 'conversion') {
            const convert = f.options.convert;
            f.options.convert = async (input, output) => {
                if (input.endsWith('beta.pptx')) throw new Error('Conversion failed on the second deck');
                return convert(input, output);
            };
        }
        if (failure === 'invalid-output') f.options.convert = async (_, output) => {
            await mkdir(output, { recursive: true });
            await writeFile(path.join(output, 'index.html'), 'not a slideshow');
        };
        if (failure === 'source-in-output') {
            const convert = f.options.convert;
            f.options.convert = async (input, output) => {
                await convert(input, output);
                await writeFile(path.join(output, 'source.pptx'), 'must not publish');
            };
        }
        await assert.rejects(updateDecks(f.options));
        assert.deepEqual(await snapshot(f.root), before);
        assert.deepEqual(await readdir(f.temporary), []);
        if (failure !== 'upload') assert.ok(!f.calls.some(call => call[0] === 'upload'));
    });
}

for (const missing of ['download-assets', 'download-controls', 'font-assets']) {
    test(`backfills ${missing} when the original source has not changed`, async t => {
        const f = await fixture(t, [published('alpha')]);
        if (missing === 'download-assets') f.options.downloads.list = async () => new Set();
        else await writeFile(path.join(f.root, 'decks', 'alpha', 'index.html'), missing === 'font-assets' ?
            addDownloadControls(html.replace('id="pp-fonts"', 'id="old-fonts"'), 'alpha') : html);
        assert.deepEqual((await updateDecks(f.options)).updated, ['alpha']);
        assert.deepEqual(f.calls.find(call => call[0] === 'upload')[1], ['alpha']);
    });
}

test('refuses to put temporary original decks anywhere inside the repository', async t => {
    const f = await fixture(t);
    f.options.tempRoot = f.root;
    await assert.rejects(updateDecks(f.options), /outside the repository/);
});

test('generated index escapes original titles', () => {
    const page = renderIndex([{ name: 'alpha', title: '<Deck> & "Notes"' }]);
    assert.match(page, /href="https:\/\/segunakinyemi\.com\/slides\/alpha\/"/);
    assert.match(page, /&lt;Deck&gt; &amp; &quot;Notes&quot;/);
    assert.doesNotMatch(page, /<Deck>/);
    assert.match(page, /Download PDF/);
    assert.match(page, /Download PowerPoint/);
    assert.match(page, /releases\/download\/downloads\/alpha\.pdf/);
});

test('enforces exact generated-file and published-site size thresholds', async t => {
    const directory = await mkdtemp(path.join(tmpdir(), 'slides-size-test-'));
    t.after(() => rm(directory, { recursive: true, force: true }));
    await writeFile(path.join(directory, 'index.html'), '12345');
    assert.equal(MAX_FILE_BYTES, 104857600);
    assert.equal(MAX_SITE_BYTES, 1_000_000_000);
    assert.equal(await validatePublicFiles(directory, { maxFileBytes: 5, maxSiteBytes: 5 }), 5);
    await assert.rejects(validatePublicFiles(directory, { maxFileBytes: 4 }), /100 MiB/);
    await assert.rejects(validatePublicFiles(directory, { maxSiteBytes: 4 }), /1 GB/);
});

// A tiny, original OOXML package exercises the real converter without retaining a PPTX fixture in Git.
function zip(parts) {
    const local = [];
    const central = [];
    let offset = 0;
    for (const [file, content] of Object.entries(parts)) {
        const name = Buffer.from(file);
        const data = Buffer.isBuffer(content) ? content : Buffer.from(content);
        const header = Buffer.alloc(30);
        header.writeUInt32LE(0x04034b50);
        header.writeUInt16LE(20, 4);
        header.writeUInt32LE(crc32(data), 14);
        header.writeUInt32LE(data.length, 18);
        header.writeUInt32LE(data.length, 22);
        header.writeUInt16LE(name.length, 26);
        const entry = Buffer.alloc(46);
        entry.writeUInt32LE(0x02014b50);
        entry.writeUInt16LE(20, 4);
        entry.writeUInt16LE(20, 6);
        entry.writeUInt32LE(crc32(data), 16);
        entry.writeUInt32LE(data.length, 20);
        entry.writeUInt32LE(data.length, 24);
        entry.writeUInt16LE(name.length, 28);
        entry.writeUInt32LE(offset, 42);
        local.push(header, name, data);
        central.push(entry, name);
        offset += header.length + name.length + data.length;
    }
    const directory = Buffer.concat(central);
    const end = Buffer.alloc(22);
    end.writeUInt32LE(0x06054b50);
    end.writeUInt16LE(Object.keys(parts).length, 8);
    end.writeUInt16LE(Object.keys(parts).length, 10);
    end.writeUInt32LE(directory.length, 12);
    end.writeUInt32LE(offset, 16);
    return Buffer.concat([...local, directory, end]);
}

const animatedGif = Buffer.concat([
    Buffer.from('GIF89a', 'ascii'),
    Buffer.from([1, 0, 1, 0, 0x80, 0, 0, 0, 0, 0, 255, 255, 255]),
    Buffer.from([0x21, 0xff, 11]), Buffer.from('NETSCAPE2.0', 'ascii'), Buffer.from([3, 1, 0, 0, 0]),
    ...[0x44, 0x4c].map(pixel => Buffer.from([
        0x21, 0xf9, 4, 0, 10, 0, 0, 0,
        0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 2, pixel, 1, 0
    ])),
    Buffer.from([0x3b])
]);

function sampleDeck() {
    const rels = 'http://schemas.openxmlformats.org/package/2006/relationships';
    const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const parts = {
        '[Content_Types].xml': `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="gif" ContentType="image/gif"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`,
        '_rels/.rels': `<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
        'ppt/presentation.xml': `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="${office}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`,
        'ppt/_rels/presentation.xml.rels': `<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${office}/slide" Target="slides/slide1.xml"/></Relationships>`,
        'ppt/slides/_rels/slide1.xml.rels': `<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${office}/image" Target="../media/image1.gif"/></Relationships>`,
        'ppt/media/image1.gif': animatedGif,
        'ppt/slides/slide1.xml': `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${office}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Reveal"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2400"/><a:t>Local converter check</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="3" name="GIF"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:transition spd="med"><p:fade/></p:transition><p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" presetID="1" presetClass="entr" nodeType="clickEffect" fill="hold"><p:childTnLst><p:set><p:cBhvr><p:cTn id="5" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing></p:sld>`
    };
    const motion = `<p:animMotion origin="layout" path="M 0 3.7037E-7 L 0 0.03542" pathEditMode="relative"><p:cBhvr><p:cTn id="6" dur="50" spd="-100000" fill="hold"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`;
    parts['ppt/slides/slide1.xml'] = parts['ppt/slides/slide1.xml'].replace('</p:set>', '</p:set>' + motion);
    return zip(parts);
}

test('real pinned PPTX2HTML converts an original fixture with player, click sequence, transition, and separate GIF', async t => {
    const directory = await mkdtemp(path.join(tmpdir(), 'slides-converter-test-'));
    t.after(() => rm(directory, { recursive: true, force: true }));
    const input = path.join(directory, 'source.pptx');
    const output = path.join(directory, 'website');
    await writeFile(input, sampleDeck());
    await convertDeck(input, output, () => {});
    const page = await readFile(path.join(output, 'index.html'), 'utf8');
    assert.match(page, /Local converter check/);
    assert.match(page, /PPTX2HTML/);
    assert.match(page, /window\.pptxPlayer/);
    assert.match(page, /assets\/image1\.gif/);
    const data = JSON.parse(page.match(/<script type="application\/json" id="pp-data">([\s\S]*?)<\/script>/)[1]);
    assert.equal(data.slides.length, 1);
    assert.ok(data.slides[0].timing.main.length > 0);
    const motion = data.slides[0].timing.main.flatMap(step => step.behaviours).find(b => b.type === 'motion');
    assert.deepEqual(motion.path, [['M', 0, 3.7037e-7], ['L', 0, 0.03542]]);
    assert.equal(motion.spd, -1);
    assert.ok(data.slides[0].timing.main[0].dur >= 50);
    assert.ok(data.slides[0].transition);
    assert.deepEqual(await readFile(path.join(output, 'assets', 'image1.gif')), animatedGif);
    await validatePublicFiles(output);
});

test('real conversion runs through the complete updater without retaining its PowerPoint source', async t => {
    const f = await fixture(t, [], [['alpha', 'alpha']]);
    f.options.convert = convertDeck;
    f.options.verify = (output, pdf, options) => verifyAndExport(output, pdf, {
        ...options,
        fontRequest: async url => new Response(url.endsWith('METADATA.pb') ? 'fonts {\nstyle: "normal"\nweight: 400\n}' :
            url.includes('fonts.googleapis.com') ?
            "@font-face{font-family:'Fixture';font-style:normal;font-weight:400;src:url(https://fonts.gstatic.com/fixture.ttf)}" :
            url.includes('fonts.gstatic.com') ? await readFile(new URL('../node_modules/pdfjs-dist/standard_fonts/LiberationSans-Regular.ttf', import.meta.url)) :
                'SIL OPEN FONT LICENSE fixture')
    });
    f.options.drive.download = async (_, destination) => writeFile(destination, sampleDeck());
    f.options.downloads.publish = async files => {
        assert.deepEqual(await readFile(files[0].pptx), sampleDeck());
        assert.equal((await readFile(files[0].pdf)).subarray(0, 5).toString(), '%PDF-');
        f.assets.add('alpha.pptx');
        f.assets.add('alpha.pdf');
        f.calls.push(['upload', files.map(file => file.name)]);
    };
    assert.deepEqual(await updateDecks(f.options), { changed: true, updated: ['alpha'], removed: [] });
    const page = await readFile(path.join(f.root, 'decks', 'alpha', 'index.html'), 'utf8');
    assert.match(page, /Local converter check/);
    assert.match(page, /assets\/image1\.gif/);
    const publication = await snapshot(path.join(f.root, 'decks'));
    assert.equal(publication['index.html'], renderIndex([{ name: 'alpha', title: 'alpha' }]));
    assert.equal(publication['slides.txt'], undefined);
    assert.doesNotMatch(JSON.stringify(publication), /source\.pptx|temporary pptx|access_token|refresh_token/);
    assert.deepEqual(await readdir(f.temporary), []);
    assert.deepEqual(await updateDecks(f.options), { changed: false, updated: [], removed: [] });
    assert.deepEqual(await snapshot(path.join(f.root, 'decks')), publication);
});
