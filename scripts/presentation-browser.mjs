import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { chromium } from 'playwright';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { correctedEngine } from './converter-player.mjs';

const { evaluate, stepEnd } = await correctedEngine();

function textKey(value) {
    return value.normalize('NFKC').replace(/[\s\u200b-\u200d\ufeff]/g, '');
}

export async function validatePdf(file, slides, width, height) {
    const loading = getDocument({ data: new Uint8Array(await readFile(file)), verbosity: 0 });
    try {
        const document = await loading.promise;
        assert.equal(document.numPages, slides.length, 'The PDF must contain exactly one page per slide.');
        for (let index = 0; index < slides.length; index++) {
            const page = await document.getPage(index + 1);
            const viewport = page.getViewport({ scale: 1 });
            assert.ok(Math.abs(viewport.width - width * 0.75) < 1 && Math.abs(viewport.height - height * 0.75) < 1,
                `PDF page ${index + 1} does not match the slide dimensions.`);
            const content = await page.getTextContent();
            const text = textKey(content.items.map(item => item.str || '').join(''));
            for (const expected of slides[index]) {
                assert.ok(text.includes(textKey(expected)), `PDF page ${index + 1} is missing revealed text: ${expected.slice(0, 90)}`);
            }
        }
    } finally {
        await loading.destroy();
    }
}

async function checkState(page, slide, step, timing, metrics, width, height) {
    const expected = evaluate(timing, { step }, key => metrics[key]?.geo, key => metrics[key]?.colors,
        key => metrics[key]?.count || 0, { aspect: width / height });
    await page.evaluate(({ slide, states }) => {
        const player = window.pptxPlayer;
        const targets = player.targetsFor(slide);
        const fail = message => { throw new Error(`Slide ${slide + 1}: ${message}`); };
        for (const [key, state] of states) {
            const entry = targets.map.get(key);
            if (!entry) fail(`animation target ${key} was not rendered`);
            for (const element of entry.els) {
                if (state.visible !== undefined && (element.style.visibility === 'hidden') !== !state.visible) {
                    fail(`reveal target ${key} has the wrong visibility`);
                }
                for (const span of element.querySelectorAll('.pp-p span')) {
                    const original = window.__slideCheck.get(span);
                    if (!original || span.classList.contains('pp-bu-gap')) continue;
                    const actual = getComputedStyle(span);
                    const wanted = { ...original.styles };
                    if (state.color) {
                        const c = state.color;
                        const color = document.createElement('span').style;
                        color.color = `rgba(${Math.round(c.r)},${Math.round(c.g)},${Math.round(c.b)},${c.a ?? 1})`;
                        wanted.color = color.color;
                        wanted.webkitTextFillColor = color.color;
                    }
                    if (state.bold !== undefined) wanted.fontWeight = state.bold ? '700' : '400';
                    if (state.italic !== undefined) wanted.fontStyle = state.italic ? 'italic' : 'normal';
                    if (state.underline !== undefined) wanted.textDecorationLine = state.underline ? 'underline' : 'none';
                    for (const property of Object.keys(wanted)) {
                        if (actual[property] !== wanted[property]) fail(`${property} changed unexpectedly on animation target ${key}`);
                    }
                }
            }
        }
    }, { slide, states: [...expected.states] });
}

export async function verifyAndExport(directory, pdf, { log = console.log } = {}) {
    const browser = await chromium.launch();
    try {
        const page = await browser.newPage();
        const errors = [];
        page.on('pageerror', error => errors.push(error.message));
        page.on('requestfailed', request => errors.push(`Asset load failed: ${request.url()}`));
        await page.addInitScript(() => {
            document.addEventListener('DOMContentLoaded', () => {
                const properties = ['color', 'webkitTextFillColor', 'fontWeight', 'fontStyle', 'textDecorationLine', 'backgroundImage'];
                window.__slideCheck = new Map([...document.querySelectorAll('.pp-p span')].map(span => {
                    const style = getComputedStyle(span);
                    const box = span.getBoundingClientRect();
                    const slide = span.closest('.pp-slide').getBoundingClientRect();
                    return [span, {
                        styles: Object.fromEntries(properties.map(property => [property, style[property]])),
                        inside: box.width > 0 && box.height > 0 && box.right > slide.left &&
                            box.left < slide.right && box.bottom > slide.top && box.top < slide.bottom
                    }];
                }));
            }, { once: true });
        });
        await page.goto(`${pathToFileURL(path.resolve(directory, 'index.html')).href}?notransitions`, { waitUntil: 'load' });
        assert.deepEqual(errors, [], 'The converted presentation failed to initialize.');
        await page.waitForFunction(() => !!window.pptxPlayer);
        await page.evaluate(async () => {
            await document.fonts.ready;
            await Promise.all([...document.images].map(image => image.decode()));
            window.pptxPlayer.silent = true;
            window.pptxPlayer.scheduleAdvance = () => {};
        });
        const data = await page.evaluate(() => JSON.parse(document.getElementById('pp-data').textContent));
        assert.ok(data.slides.length > 0, 'The presentation has no slides.');
        await page.setViewportSize({ width: Math.ceil(data.width), height: Math.ceil(data.height) });
        const finalSlides = [];
        for (let slide = 0; slide < data.slides.length; slide++) {
            const timing = data.slides[slide].timing;
            const metrics = await page.evaluate(index => {
                const player = window.pptxPlayer;
                player.seek(index, 0);
                player.stepStart = null;
                const targets = player.targetsFor(index);
                return Object.fromEntries([...targets.map.keys()].map(key => [key, {
                    geo: targets.geo(key), colors: targets.colors(key), count: targets.counts.get(key) || 0
                }]));
            }, slide);
            const steps = timing?.main || [];
            for (let step = 0; step <= steps.length; step++) {
                await page.evaluate(({ slide, step }) => {
                    const player = window.pptxPlayer;
                    player.seek(slide, step);
                    player.stepStart = null;
                }, { slide, step });
                await checkState(page, slide, step, timing, metrics, data.width, data.height);
                const duration = step < steps.length ? stepEnd(steps[step], key => metrics[key]?.count || 0) : 0;
                if (step < steps.length && !steps[step].auto && Number.isFinite(duration) && duration <= 10_000) {
                    await page.getByRole('button', { name: 'Next', exact: true }).click({ force: true });
                    await page.waitForFunction(({ slide, step }) => {
                        const state = window.pptxPlayer.state();
                        return state.slide === slide && state.step >= step + 1 && !state.running;
                    }, { slide, step }, { timeout: Math.max(5_000, duration + 3_000) });
                    const played = await page.evaluate(() => window.pptxPlayer.state().step);
                    await checkState(page, slide, played, timing, metrics, data.width, data.height);
                }
            }
            if (steps.length) {
                await page.evaluate(({ slide, steps }) => {
                    const player = window.pptxPlayer;
                    player.seek(slide, steps);
                    player.stepStart = null;
                    player.prev();
                }, { slide, steps: steps.length });
                await checkState(page, slide, steps.length - 1, timing, metrics, data.width, data.height);
                await page.evaluate(index => {
                    const player = window.pptxPlayer;
                    player.seek(index, 0);
                    player.stepStart = null;
                }, slide);
                await checkState(page, slide, 0, timing, metrics, data.width, data.height);
            }
            const completed = await page.evaluate(index => {
                const player = window.pptxPlayer;
                const timing = player.data.slides[index].timing;
                player.seek(index, timing?.main?.length || 0);
                player.stepStart = null;
                player.triggers = (timing?.triggers || []).map((sequence, i) => ({ index: i, step: sequence.steps.length, start: null }));
                player.render(performance.now());
                cancelAnimationFrame(player.raf);
                clearTimeout(player.advTimer);
                const section = player.sections[index];
                const slideBox = section.getBoundingClientRect();
                const text = [...section.querySelectorAll('.pp-p span')].filter(span => {
                    const style = getComputedStyle(span);
                    const box = span.getBoundingClientRect();
                    const inside = box.width > 0 && box.height > 0 && box.right > slideBox.left &&
                        box.left < slideBox.right && box.bottom > slideBox.top && box.top < slideBox.bottom;
                    // Authored bounds keep a broken animation from hiding missing text; off-slide template notes are not PDF content.
                    const belongs = window.__slideCheck.get(span)?.inside || inside;
                    return style.visibility !== 'hidden' && Number(style.opacity) > 0 &&
                        belongs && !span.closest('.pp-bu') && /[\p{L}\p{N}]/u.test(span.textContent);
                }).map(span => span.textContent);
                const clone = section.cloneNode(true);
                clone.classList.remove('pp-current', 'pp-prev', 'pp-on-top');
                return { html: clone.outerHTML, text };
            }, slide);
            finalSlides.push(completed);
        }
        assert.deepEqual(errors, [], 'The converted presentation has browser or asset errors.');
        await page.evaluate(({ slides, width, height }) => {
            window.pptxPlayer.destroy();
            document.documentElement.classList.remove('pp-deck-mode');
            document.body.innerHTML = `<main class="slide-print">${slides.join('\n')}</main>`;
            const style = document.createElement('style');
            style.textContent = `
@page{size:${width}px ${height}px;margin:0}
html,body{margin:0;padding:0;width:${width}px;height:auto;overflow:visible;background:white}
*{-webkit-print-color-adjust:exact;print-color-adjust:exact}
.slide-print>.pp-slide{display:block;position:relative;width:${width}px;height:${height}px;margin:0;box-shadow:none;break-inside:avoid;break-after:page}
.slide-print>.pp-slide:last-child{break-after:auto}
.pp-ui,.pp-progress,.pp-notes,.pp-blank{display:none!important}`;
            document.head.appendChild(style);
        }, { slides: finalSlides.map(slide => slide.html), width: data.width, height: data.height });
        await page.evaluate(async () => {
            await document.fonts.ready;
            await Promise.all([...document.images].map(image => image.decode()));
        });
        await page.pdf({ path: pdf, printBackground: true, preferCSSPageSize: true, displayHeaderFooter: false });
        assert.deepEqual(errors, [], 'The PDF export has browser or asset errors.');
        await validatePdf(pdf, finalSlides.map(slide => slide.text), data.width, data.height);
        log(`Verified browser playback and PDF (${data.slides.length} slides/pages).`);
        return { slides: data.slides.length };
    } finally {
        await browser.close();
    }
}
