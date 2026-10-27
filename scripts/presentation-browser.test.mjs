import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { test } from 'node:test';
import { BASE_CSS } from 'pptx2html';
import { chromium } from 'playwright';
import { converterRoot, correctedPlayer, correctPlayerSource, correctTimingSource, correctEngineSource, correctStartupSource, correctedEngine } from './converter-player.mjs';
import { addDownloadControls } from './downloads.mjs';
import { verifyAndExport, validatePdf, verifyMedia } from './presentation-browser.mjs';

// Original two-second 16x16 red H.264 test clip, generated with FFmpeg.
const video = Buffer.from(
    'AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDEAAAMwbW9vdgAAAGxtdmhkAAAAAAAAAAAAAAAAAAAD6AAAB9AAAQAAAQAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAlt0cmFrAAAAXHRraGQAAAADAAAAAAAAAAAAAAABAAAAAAAAB9AAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAABAAAAAQAAAAAAAkZWR0cwAAABxlbHN0AAAAAAAAAAEAAAfQAAAAAAABAAAAAAHTbWRpYQAAACBtZGhkAAAAAAAAAAAAAAAAAABAAAAAgABVxAAAAAAALWhkbHIAAAAAAAAAAHZpZGUAAAAAAAAAAAAAAABWaWRlb0hhbmRsZXIAAAABfm1pbmYAAAAUdm1oZAAAAAEAAAAAAAAAAAAAACRkaW5mAAAAHGRyZWYAAAAAAAAAAQAAAAx1cmwgAAAAAQAAAT5zdGJsAAAAtnN0c2QAAAAAAAAAAQAAAKZhdmMxAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAABAAEABIAAAASAAAAAAAAAABFExhdmM2MC45LjEwMCBsaWJ4MjY0AAAAAAAAAAAAAAAAGP//AAAALGF2Y0MBQsAK/+EAFWdCwAraewEQAAADABAAAAMAQPEiagEABGjOD8gAAAAQcGFzcAAAAAEAAAABAAAAFGJ0cnQAAAAAAAAKJAAACiQAAAAYc3R0cwAAAAAAAAABAAAABAAAIAAAAAAUc3NzcwAAAAAAAAABAAAAAQAAABxzdHNjAAAAAAAAAAEAAAABAAAABAAAAAEAAAAkc3RzegAAAAAAAAAAAAAABAAAAm4AAAAJAAAACQAAAAkAAAAUc3RjbwAAAAAAAAABAAADYAAAAGF1ZHRhAAAAWW1ldGEAAAAAAAAAIWhkbHIAAAAAAAAAAG1kaXJhcHBsAAAAAAAAAAAAAAAALGlsc3QAAAAkqXRvbwAAABxkYXRhAAAAAQAAAABMYXZmNjAuNC4xMDEAAAAIZnJlZQAAApFtZGF0AAACUwYF//9P3EXpvebZSLeWLNgg2SPu73gyNjQgLSBjb3JlIDE2NCByMzEwNiBlYWE2OGZhIC0gSC4yNjQvTVBFRy00IEFWQyBjb2RlYyAtIENvcHlsZWZ0IDIwMDMtMjAyMyAtIGh0dHA6Ly93d3cudmlkZW9sYW4ub3JnL3gyNjQuaHRtbCAtIG9wdGlvbnM6IGNhYmFjPTAgcmVmPTEgZGVibG9jaz0wOjA6MCBhbmFseXNlPTA6MCBtZT1kaWEgc3VibWU9MCBwc3k9MSBwc3lfcmQ9MS4wMDowLjAwIG1peGVkX3JlZj0wIG1lX3JhbmdlPTE2IGNocm9tYV9tZT0xIHRyZWxsaXM9MCA4eDhkY3Q9MCBjcW09MCBkZWF6b25lPTIxLDExIGZhc3RfcHNraXA9MSBjaHJvbWFfcXBfb2Zmc2V0PTAgdGhyZWFkcz0xIGxvb2thaGVhZF90aHJlYWRzPTEgc2xpY2VkX3RocmVhZHM9MCBucj0wIGRlY2ltYXRlPTEgaW50ZXJsYWNlZD0wIGJsdXJheV9jb21wYXQ9MCBjb25zdHJhaW5lZF9pbnRyYT0wIGJmcmFtZXM9MCB3ZWlnaHRwPTAga2V5aW50PTI1MCBrZXlpbnRfbWluPTIgc2NlbmVjdXQ9MCBpbnRyYV9yZWZyZXNoPTAgcmM9Y3JmIG1idHJlZT0wIGNyZj0yMy4wIHFjb21wPTAuNjAgcXBtaW49MCBxcG1heD02OSBxcHN0ZXA9NCBpcF9yYXRpbz0xLjQwIGFxPTAgAIAAAAATZYiEOhGKAAIY8cAAQPY4AAh5YAAAAAVBmiASpQAAAAVBmkATpQAAAAVBmmATpQ==',
    'base64');

async function fixture(t, player, { missingImage = false, media = false } = {}) {
    const root = await mkdtemp(path.join(tmpdir(), 'slides-format-test-'));
    t.after(() => rm(root, { recursive: true, force: true }));
    const behaviour = (spid, effect) => ({
        type: 'set', attrs: ['style.visibility'], to: 'visible', spid, keys: [spid],
        dur: 1, repeat: 1, spd: 1, fill: 'hold', begin: 0, effect, cls: 'entr', wrappers: []
    });
    const steps = ['2', '3'].map((id, index) => ({
        auto: false, delay: 0, dur: 50, behaviours: [
            behaviour(id, index),
            { ...behaviour(id, index), type: 'effect', attrs: [], filter: 'fade', transition: 'in', dur: 50, fill: 'remove' }
        ]
    }));
    const data = { width: 640, height: 360, slides: [
        { index: 0, number: 1, hidden: false, timing: { main: steps, triggers: [], hidden: ['2', '3'] } },
        { index: 1, number: 2, hidden: false, timing: null }
    ] };
    const html = `<!doctype html><html><head><style>${BASE_CSS}</style></head><body class="pp-body">
<div class="pp-deck pp-deck" data-width="640" data-height="360">
<section class="pp-slide" id="slide-1" style="width:640px;height:360px">
<div class="pp-bg" style="background:#000"></div>
<div class="pp-el" data-spid="2" data-layer="slide" style="left:30px;top:30px;width:550px;height:100px">
<div class="pp-p"><span style="font-size:24px;color:#f8fafc;font-weight:bold;font-style:italic;text-decoration-line:underline">First reveal </span><span style="font-size:24px;color:#fff;background:linear-gradient(white,white);background-clip:text;-webkit-text-fill-color:transparent">Gradient</span></div></div>
<div class="pp-el" data-spid="3" data-layer="slide" style="left:30px;top:160px;width:550px;height:100px">
<div class="pp-p"><span style="font-size:24px;color:#50e6ff">Second reveal</span></div></div>
${missingImage ? '<img src="missing.png">' : ''}
${media ? '<video src="clip.mp4" preload="metadata" width="16" height="16"></video>' : ''}
</section>
<section class="pp-slide" id="slide-2" style="width:640px;height:360px"><div class="pp-p"><span style="font-size:24px;color:#000">Last slide</span></div><div class="pp-el" style="left:700px;top:10px;width:200px;height:40px"><div class="pp-p"><span style="font-size:24px">Off-slide authoring note</span></div></div></section>
</div>
<script type="application/json" id="pp-data">${JSON.stringify(data)}</script>
<script>${player}</script></body></html>`;
    await writeFile(path.join(root, 'index.html'), html);
    if (media) await writeFile(path.join(root, 'clip.mp4'), video);
    return root;
}

test('a valid MP4 decodes, plays, seeks and exports without treating metadata cancellation as a broken asset', async t => {
    const root = await fixture(t, await correctedPlayer(), { media: true });
    const messages = [];
    assert.deepEqual(await verifyAndExport(root, path.join(root, 'reading-copy.pdf'), {
        log: message => messages.push(message)
    }), { slides: 2 });
    assert.ok(messages.some(message => /Verified playback and seeking for 1 media assets/.test(message)));
});

for (const failure of ['missing', 'unsupported']) {
    test(`${failure} MP4 files still stop publication`, async t => {
        const root = await fixture(t, await correctedPlayer(), { media: true });
        if (failure === 'missing') await rm(path.join(root, 'clip.mp4'));
        else await writeFile(path.join(root, 'clip.mp4'), 'not a supported video');
        await assert.rejects(verifyAndExport(root, path.join(root, 'unpublished.pdf'), { log: () => {} }),
            /Asset load failed|Media .*clip\.mp4/);
    });
}

test('a release media URL remains a playable, seekable subresource and failed playback is not ignored', async t => {
    const browser = await chromium.launch();
    t.after(() => browser.close());
    const page = await browser.newPage();
    const requests = [];
    const url = 'https://github.com/segunak/slides/releases/download/downloads/alpha--media--media1.mp4';
    await page.route(url, route => {
        requests.push(route.request().url());
        return route.fulfill({ contentType: 'video/mp4',
            headers: { 'Content-Disposition': 'attachment; filename="alpha--media--media1.mp4"', 'Accept-Ranges': 'bytes' },
            body: video });
    });
    await page.setContent(`<video src="${url}" preload="metadata"></video>`);
    assert.equal(await verifyMedia(page), 1);
    assert.ok(requests.length > 0 && requests.every(request => request === url));
    await page.locator('video').evaluate(async media => {
        media.muted = true;
        const ended = new Promise(resolve => media.addEventListener('ended', resolve, { once: true }));
        await media.play();
        await ended;
    });
    assert.equal(await page.locator('video').evaluate(media => media.ended), true);
    assert.equal(await verifyMedia(page), 1, 'Previously completed media must still pass playback verification.');
    await page.locator('video').evaluate(media => {
        media.play = () => Promise.reject(new Error('Playback failed'));
    });
    await assert.rejects(verifyMedia(page), /Playback failed/);
});

test('the style correction refuses unknown converter source instead of patching it blindly', () => {
    for (const correct of [correctPlayerSource, correctTimingSource, correctEngineSource, correctStartupSource]) {
        assert.throws(() => correct('unrelated source'), /pinned PPTX2HTML source changed/);
    }
});

test('player layout and download controls wait for fonts before initialization', async t => {
    const root = await fixture(t, await correctedPlayer());
    const file = path.join(root, 'index.html');
    const html = (await readFile(file, 'utf8')).replace('<head>',
        '<head><script>window.pptxFontsReady=new Promise(resolve=>window.releaseFonts=resolve)</script>');
    await writeFile(file, addDownloadControls(html, 'alpha'));
    const browser = await chromium.launch();
    t.after(() => browser.close());
    const page = await browser.newPage();
    await page.goto(pathToFileURL(file).href);
    assert.equal(await page.evaluate(() => !!window.pptxPlayer), false);
    await page.evaluate(() => window.releaseFonts());
    await page.waitForFunction(() => !!window.pptxPlayer);
    assert.equal(await page.getByRole('link', { name: 'Download PDF', exact: true }).count(), 1);
});

test('signed animation speed changes direction without making duration negative', async () => {
    const { activeDuration, behaviourProgress, pathPoint, stepEnd } = await correctedEngine();
    const motion = { type: 'motion', begin: 100, dur: 700, spd: -1, fill: 'hold', repeat: 1 };
    assert.equal(activeDuration(motion), 700);
    assert.equal(stepEnd({ behaviours: [motion] }), 800);
    assert.deepEqual(behaviourProgress(motion, 100), { p: 1, running: true });
    assert.deepEqual(behaviourProgress(motion, 450), { p: 0.5, running: true });
    assert.deepEqual(behaviourProgress(motion, 800), { p: 0, running: false });
    const points = [['M', 0, 3.7037e-7], ['L', 0, 0.03542]];
    assert.deepEqual(pathPoint(points, behaviourProgress(motion, 800).p), [0, 3.7037e-7]);
    assert.deepEqual(behaviourProgress({ ...motion, spd: 1 }, 800), { p: 1, running: false });
    assert.deepEqual(behaviourProgress({ ...motion, spd: -2 }, 450), { p: 0, running: false });
});

test('browser verification catches the original black-text bug', async t => {
    const { PLAYER_SOURCE } = await import(pathToFileURL(path.join(converterRoot, 'src', 'generated', 'player-source.js')));
    const root = await fixture(t, PLAYER_SOURCE);
    await assert.rejects(verifyAndExport(root, path.join(root, 'unpublished.pdf'), { log: () => {} }),
        /color changed unexpectedly|webkitTextFillColor changed unexpectedly/);
});

test('corrected reveals preserve white, cyan, gradient, bold, italic and underline; PDF has one completed slide per page', async t => {
    const root = await fixture(t, await correctedPlayer());
    const pdf = path.join(root, 'reading-copy.pdf');
    assert.deepEqual(await verifyAndExport(root, pdf, { log: () => {} }), { slides: 2 });
    await validatePdf(pdf, [['First reveal', 'Gradient', 'Second reveal'], ['Last slide']], 640, 360);
    await assert.rejects(validatePdf(pdf, [['First reveal']], 640, 360), /exactly one page/);
    await assert.rejects(validatePdf(pdf, [[], []], 1280, 720), /slide dimensions/);
    await assert.rejects(validatePdf(pdf, [['Text not present'], []], 640, 360), /missing revealed text/);
    assert.equal((await readFile(pdf)).subarray(0, 5).toString(), '%PDF-');
});

test('a missing local media asset prevents publication instead of silently producing an incomplete PDF', async t => {
    const root = await fixture(t, await correctedPlayer(), { missingImage: true });
    await assert.rejects(verifyAndExport(root, path.join(root, 'unpublished.pdf'), { log: () => {} }),
        /failed to initialize|browser or asset errors|source image cannot be decoded/i);
});

test('moving authored slide text off-canvas still fails verification instead of excluding it from PDF expectations', async t => {
    const broken = `${await correctedPlayer()}
document.addEventListener('DOMContentLoaded',()=>{
  const p=window.pptxPlayer, render=p.render.bind(p);
  p.render=(...args)=>{
    const result=render(...args);
    p.targetsFor(0).map.get('2').els[0].style.transform='translateY(2000px)';
    return result;
  };
},{once:true});`;
    const root = await fixture(t, broken);
    await assert.rejects(verifyAndExport(root, path.join(root, 'unpublished.pdf'), { log: () => {} }), /missing revealed text/);
});

test('download controls initialize, fit a narrow screen, and download without advancing the presentation', async t => {
    const root = await fixture(t, await correctedPlayer());
    const html = addDownloadControls(await readFile(path.join(root, 'index.html'), 'utf8'), 'alpha');
    const browser = await chromium.launch();
    t.after(() => browser.close());
    const page = await browser.newPage({ viewport: { width: 390, height: 844 } });
    await page.route('https://github.com/segunak/slides/releases/download/downloads/*', route => route.fulfill({
        headers: { 'Content-Disposition': 'attachment; filename="alpha.pdf"' },
        contentType: 'application/pdf',
        body: '%PDF-fixture'
    }));
    await page.setContent(html, { waitUntil: 'load' });
    const before = await page.evaluate(() => window.pptxPlayer.state());
    for (const [label, extension] of [['Download PDF', 'pdf'], ['Download PowerPoint', 'pptx']]) {
        const link = page.getByRole('link', { name: label, exact: true });
        assert.equal(await link.getAttribute('href'), `https://github.com/segunak/slides/releases/download/downloads/alpha.${extension}`);
        const bounds = await link.boundingBox();
        assert.ok(bounds && bounds.x >= 0 && bounds.x + bounds.width <= 390, 'Download links must fit the mobile viewport.');
        const completed = page.waitForEvent('download');
        await link.click();
        await completed;
        const after = await page.evaluate(() => window.pptxPlayer.state());
        assert.equal(after.slide, before.slide);
        assert.equal(after.step, before.step);
    }
});
