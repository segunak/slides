import assert from 'node:assert/strict';
import { mkdtemp, readFile, readdir, rm, unlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { test } from 'node:test';
import { chromium } from 'playwright';
import { bundleFonts, fontChoice, verifyFonts } from './fonts.mjs';

async function fixture(t) {
    const root = await mkdtemp(path.join(tmpdir(), 'slides-font-test-'));
    t.after(() => rm(root, { recursive: true, force: true }));
    const font = await readFile(new URL('../node_modules/pdfjs-dist/standard_fonts/LiberationSans-Regular.ttf', import.meta.url));
    const requests = [];
    const request = async url => {
        requests.push(url);
        if (url.endsWith('METADATA.pb')) return new Response('fonts {\nstyle: "normal"\nweight: 400\n}\nfonts {\nstyle: "italic"\nweight: 700\n}\naxes {\ntag: "wght"\nmin_value: 100\nmax_value: 900\n}');
        if (url.includes('fonts.gstatic.com')) return new Response(font);
        if (url.includes('fonts.googleapis.com')) return new Response(
            "@font-face{font-family:'Fixture';font-style:normal;font-weight:400;src:url(https://fonts.gstatic.com/fixture.ttf)}");
        return new Response('SIL OPEN FONT LICENSE fixture');
    };
    const browser = await chromium.launch();
    t.after(() => browser.close());
    await writeFile(path.join(root, 'index.html'), `<!doctype html><html><head></head><body>
<section class="pp-slide" id="slide-1">
<span id="original" style="font-family:Arial;font-size:24px">Installed original</span>
<span id="fallback" style="font-family:'Imaginary Sans';font-size:24px">Unknown original</span>
<span id="italic" style="font-family:'Imaginary Sans';font-size:24px;font-weight:bold;font-style:italic">Bold italic</span>
<svg><text font-family="Imaginary Mono">Chart label</text></svg>
</section></body></html>`);
    return { root, browser, requests, request };
}

test('uses upstream compatible substitutes and category classification without an exhaustive source-font map', () => {
    for (const [original, family] of [
        ['Arial', 'Arimo'], ['Calibri', 'Carlito'], ['Cambria', 'Caladea'], ['Times New Roman', 'Tinos'],
        ['Courier New', 'Cousine'], ['Segoe UI', 'Roboto'], ['Unknown Sans', 'Roboto'],
        ['Unknown Roman', 'Tinos'], ['Unknown Mono', 'Cousine'], ['Gelasio', 'Gelasio']
    ]) assert.equal(fontChoice(original).family, family);
    assert.equal(fontChoice('Segoe UI Semibold').downloadWeight, 600);
    assert.equal(fontChoice('Calibri Light').downloadWeight, 300);
    assert.equal(fontChoice('Unknown Sans', '700', 'italic').downloadStyle, 'italic');
    assert.equal(fontChoice('Segoe UI Bold Italic').downloadStyle, 'italic');
    assert.equal(fontChoice('Segoe UI Bold Italic').downloadWeight, 700);
});

test('bundles only requested styles, original-first CSS, local assets, notices, and SVG fonts', async t => {
    const f = await fixture(t);
    await bundleFonts(f.browser, f.root, f.request);
    const html = await readFile(path.join(f.root, 'index.html'), 'utf8');
    assert.match(html, /id="pp-fonts"/);
    assert.match(html, /pptxFontsReady/);
    assert.match(html, /font-family: Arial, pp-web-0, sans-serif/);
    assert.ok(!html.includes('https://fonts.gstatic.com'));
    assert.ok(f.requests.some(url => url.includes('Roboto') && url.includes('1%2C700')));
    assert.ok(f.requests.some(url => url.includes('Cousine')));
    const files = await readdir(path.join(f.root, 'assets', 'fonts'));
    assert.equal(files.filter(name => name.endsWith('.ttf')).length, 1, 'Identical font assets are deduplicated.');
    assert.ok(files.includes('roboto-OFL.txt'));
    const count = f.requests.length;
    await bundleFonts(f.browser, f.root, () => assert.fail('Already bundled fonts must not be downloaded again.'));
    assert.equal(f.requests.length, count);
    const page = await f.browser.newPage();
    const remoteRequests = [];
    page.on('request', request => { if (/^https?:/.test(request.url())) remoteRequests.push(request.url()); });
    await page.goto(pathToFileURL(path.join(f.root, 'index.html')).href);
    await page.evaluate(() => window.pptxFontsReady);
    const session = await page.context().newCDPSession(page);
    await session.send('DOM.enable');
    await session.send('CSS.enable');
    const { root } = await session.send('DOM.getDocument');
    const actual = async selector => {
        const { nodeId } = await session.send('DOM.querySelector', { nodeId: root.nodeId, selector });
        return (await session.send('CSS.getPlatformFontsForNode', { nodeId })).fonts;
    };
    if (process.platform === 'win32') {
        const original = await actual('#original');
        assert.ok(original.some(font => font.familyName === 'Arial' && !font.isCustomFont));
    }
    assert.ok((await actual('#fallback')).every(font => font.isCustomFont));
    assert.ok((await actual('#italic')).every(font => font.isCustomFont));
    assert.deepEqual(remoteRequests, []);
    const messages = [];
    await verifyFonts(page, session, 0, message => messages.push(message), new Set());
    assert.ok(messages.some(message => /Imaginary Sans.*\[bundled\]/.test(message)));
    for (let attempt = 0; attempt < 20; attempt++) {
        await page.locator('#fallback').evaluate(e => { e.style.fontFamily = '"Imaginary Sans", "pp-web-1", sans-serif'; });
        await verifyFonts(page, session, 0, () => {}, new Set());
        await page.locator('#fallback').evaluate(e => { e.style.fontFamily = 'monospace'; });
        await assert.rejects(verifyFonts(page, session, 0, () => {}, new Set()), /unexpected font substitution/);
    }
});

test('font download and invalid-CSS failures are explicit and leave the original HTML intact', async t => {
    const f = await fixture(t);
    const file = path.join(f.root, 'index.html');
    const before = await readFile(file, 'utf8');
    await assert.rejects(bundleFonts(f.browser, f.root, async () => new Response('unavailable', { status: 503 })), /HTTP 503/);
    await assert.rejects(bundleFonts(f.browser, f.root, async url => url.endsWith('METADATA.pb') ? f.request(url) :
        new Response('@font-face{src:url(https://unexpected.example/font)}')),
        /Unexpected Google Fonts source/);
    assert.equal(await readFile(file, 'utf8'), before);
});

test('uses Google metadata to select an available weight and style instead of maintaining font capabilities by hand', async t => {
    const f = await fixture(t);
    await writeFile(path.join(f.root, 'index.html'), `<html><head></head><body><section class="pp-slide">
<span style="font-family:Arial;font-weight:900">Heavy text</span>
<span style="font-family:'Unknown Mono';font-style:italic">Italic text</span>
</section></body></html>`);
    await bundleFonts(f.browser, f.root, async url => url.endsWith('METADATA.pb') ?
        new Response('fonts {\nstyle: "normal"\nweight: 400\n}\nfonts {\nstyle: "normal"\nweight: 700\n}') : f.request(url));
    const cssRequests = f.requests.filter(url => url.includes('fonts.googleapis.com'));
    assert.ok(cssRequests.some(url => url.includes('Arimo') && url.includes('0%2C700')));
    assert.ok(cssRequests.some(url => url.includes('Cousine') && url.includes('0%2C400')));
    assert.ok(!cssRequests.some(url => url.includes('900') || url.includes('1%2C')));
});

test('an embedded original remains ahead of the bundled fallback', async t => {
    const f = await fixture(t);
    const file = path.join(f.root, 'index.html');
    const font = await readFile(new URL('../node_modules/pdfjs-dist/standard_fonts/LiberationSans-Bold.ttf', import.meta.url));
    await writeFile(file, `<!doctype html><html><head><style>@font-face{font-family:Embedded;src:url(data:font/ttf;base64,${font.toString('base64')})}</style></head>
<body><section class="pp-slide" id="slide-1"><span style="font-family:Embedded">Embedded original</span></section></body></html>`);
    await bundleFonts(f.browser, f.root, f.request);
    const page = await f.browser.newPage();
    await page.goto(pathToFileURL(file).href);
    await page.evaluate(() => window.pptxFontsReady);
    assert.equal(await page.locator('span').evaluate(e => getComputedStyle(e).fontFamily.split(',')[0]), 'Embedded');
    const session = await page.context().newCDPSession(page);
    await session.send('DOM.enable');
    await session.send('CSS.enable');
    const { root } = await session.send('DOM.getDocument');
    const { nodeId } = await session.send('DOM.querySelector', { nodeId: root.nodeId, selector: 'span' });
    const { fonts } = await session.send('CSS.getPlatformFontsForNode', { nodeId });
    assert.ok(fonts.some(font => font.isCustomFont && /Bold/.test(font.postScriptName)));
});

test('a missing bundled font rejects readiness rather than quietly using a system fallback', async t => {
    const f = await fixture(t);
    await bundleFonts(f.browser, f.root, f.request);
    const folder = path.join(f.root, 'assets', 'fonts');
    const file = (await readdir(folder)).find(name => name.endsWith('.ttf'));
    await unlink(path.join(folder, file));
    const page = await f.browser.newPage();
    await page.goto(pathToFileURL(path.join(f.root, 'index.html')).href);
    await assert.rejects(page.evaluate(() => window.pptxFontsReady), /network|font|load/i);
});
