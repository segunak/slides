import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { converterRoot } from './converter-player.mjs';

// Read the pinned converter's actual mappings and classification, rather than maintaining another map.
const source = await readFile(path.join(converterRoot, 'src', 'render', 'fonts.js'), 'utf8');
const { SUBSTITUTES, SERIF, MONO } = await import(`data:text/javascript;base64,${Buffer.from(
    source + '\nexport { SUBSTITUTES, SERIF, MONO };').toString('base64')}`);
const googleFamilies = ['Arimo', 'Carlito', 'Caladea', 'Tinos', 'Cousine', 'Roboto',
    'Gelasio', 'EB Garamond', 'Inconsolata', 'Libre Franklin', 'STIX Two Math'];
const downloads = new Map();
export const fontMarker = 'id="pp-fonts"';
const baseFamily = name => name.replace(/(?:\s+(?:semi[ -]?bold|demi[ -]?bold|light|black|bold|italic))+$/i, '');

export function fontChoice(original, weight = '400', style = 'normal') {
    const base = baseFamily(original);
    const candidates = [original, ...(SUBSTITUTES[original.toLowerCase()] || SUBSTITUTES[base.toLowerCase()] || [])];
    const generic = MONO.test(original) ? 'monospace' : SERIF.test(original) ? 'serif' : 'sans-serif';
    const family = candidates.map(name => googleFamilies.find(f => f.toLowerCase() === name.toLowerCase())).find(Boolean) ||
        { monospace: 'Cousine', serif: 'Tinos', 'sans-serif': 'Roboto' }[generic];
    const namedWeight = /semi[ -]?bold|demi[ -]?bold/i.test(original) ? 600 :
        /\blight\b/i.test(original) ? 300 : /\bblack\b/i.test(original) ? 900 : /\bbold\b/i.test(original) ? 700 : 400;
    return { original, family, generic, weight, style,
        downloadWeight: Number(weight) === 400 ? namedWeight : Number(weight),
        downloadStyle: style !== 'normal' || /\bitalic\b/i.test(original) ? 'italic' : 'normal' };
}

async function download(url, request) {
    if (request === fetch && downloads.has(url)) return downloads.get(url);
    const response = await request(url, {
        signal: AbortSignal.timeout(30_000),
        headers: { 'User-Agent': 'Mozilla/5.0 Chrome/130.0.0.0 Safari/537.36' }
    });
    assert.ok(response.ok, `Font download failed (HTTP ${response.status}): ${url}`);
    const bytes = Buffer.from(await response.arrayBuffer());
    assert.ok(bytes.length, `Empty font response: ${url}`);
    if (request === fetch) downloads.set(url, bytes);
    return bytes;
}

export async function bundleFonts(browser, directory, request = fetch) {
    const htmlFile = path.join(directory, 'index.html');
    const html = await readFile(htmlFile, 'utf8');
    if (html.includes(fontMarker)) return;
    const page = await browser.newPage({ javaScriptEnabled: false });
    try {
        await page.goto(pathToFileURL(htmlFile).href);
        const uses = await page.evaluate(() => {
            const specs = [];
            for (const element of document.querySelectorAll('.pp-slide [style*="font-family"], .pp-slide text[font-family]')) {
                if (!element.textContent.trim()) continue;
                const css = getComputedStyle(element);
                const original = css.fontFamily.split(',')[0].trim().replace(/^['"]|['"]$/g, '');
                const value = { original, weight: css.fontWeight, style: css.fontStyle };
                let id = specs.findIndex(s => JSON.stringify(s) === JSON.stringify(value));
                if (id < 0) { id = specs.length; specs.push(value); }
                element.dataset.ppFont = String(id);
            }
            return specs;
        });
        const specs = uses.map((use, id) => ({ ...fontChoice(use.original, use.weight, use.style), alias: `pp-web-${id}` }));
        const assets = path.join(directory, 'assets', 'fonts');
        let css = '';
        for (const spec of specs) {
            const slug = spec.family.toLowerCase().replaceAll(' ', '');
            const fontRoot = `https://raw.githubusercontent.com/google/fonts/main/ofl/${slug}`;
            const metadata = (await download(`${fontRoot}/METADATA.pb`, request)).toString();
            const variants = [...metadata.matchAll(/fonts \{([^}]+)\}/g)].map(([, font]) => ({
                style: font.match(/style: "([^"]+)"/)?.[1], weight: Number(font.match(/weight: (\d+)/)?.[1])
            }));
            assert.ok(variants.length, `Missing Google font metadata for ${spec.family}.`);
            if (!variants.some(v => v.style === spec.downloadStyle)) spec.downloadStyle = 'normal';
            const axis = metadata.match(/axes \{\s*tag: "wght"\s*min_value: ([\d.]+)\s*max_value: ([\d.]+)/);
            const weights = variants.filter(v => v.style === spec.downloadStyle).map(v => v.weight);
            assert.ok(weights.length, `No usable Google font styles for ${spec.family}.`);
            spec.downloadWeight = axis ? Math.max(Number(axis[1]), Math.min(Number(axis[2]), spec.downloadWeight)) :
                weights.reduce((best, weight) => Math.abs(weight - spec.downloadWeight) < Math.abs(best - spec.downloadWeight) ? weight : best);
            const query = new URLSearchParams({
                family: `${spec.family}:ital,wght@${spec.downloadStyle === 'italic' ? 1 : 0},${spec.downloadWeight}`,
                display: 'block'
            });
            const remoteCss = (await download(`https://fonts.googleapis.com/css2?${query}`, request)).toString();
            const faces = [...remoteCss.matchAll(/@font-face\s*\{([^}]+)\}/g)];
            assert.ok(faces.length, `Google Fonts did not return ${spec.family}.`);
            await mkdir(assets, { recursive: true });
            for (const [, body] of faces) {
                const url = body.match(/src:\s*url\((https:\/\/fonts\.gstatic\.com\/[^)\s]+)\)/)?.[1];
                assert.ok(url, `Unexpected Google Fonts source for ${spec.family}.`);
                const bytes = await download(url, request);
                const ext = new URL(url).pathname.split('.').pop();
                assert.ok(['woff2', 'woff', 'ttf'].includes(ext), 'Unsupported Google font format.');
                const file = `${createHash('sha256').update(bytes).digest('hex').slice(0, 20)}.${ext}`;
                await writeFile(path.join(assets, file), bytes);
                css += `@font-face{${body.replace(/font-family:[^;]+;/, `font-family:'${spec.alias}';`)
                    .replace(/font-weight:[^;]+;/, `font-weight:${spec.weight};`)
                    .replace(/font-style:[^;]+;/, `font-style:${spec.downloadStyle === 'italic' ? spec.style : 'normal'};`)
                    .replace(url, `assets/fonts/${file}`)}}\n`;
            }
            const licenseUrl = spec.family === 'Tinos' ?
                'https://raw.githubusercontent.com/googlefonts/tinos/3b4482a99b80ea5fc75f187b1be3120a3f5905b3/OFL.txt' :
                `${fontRoot}/OFL.txt`;
            await writeFile(path.join(assets, `${slug}-OFL.txt`), await download(licenseUrl, request));
        }
        await page.evaluate(({ specs, css }) => {
            for (const element of document.querySelectorAll('[data-pp-font]')) {
                const spec = specs[Number(element.dataset.ppFont)];
                element.style.fontFamily = `${JSON.stringify(spec.original)}, "${spec.alias}", ${spec.generic}`;
            }
            const style = document.createElement('style');
            style.textContent = css;
            document.head.appendChild(style);
            const data = document.createElement('script');
            data.id = 'pp-fonts';
            data.type = 'application/json';
            data.textContent = JSON.stringify(specs).replaceAll('<', '\\u003c');
            document.head.appendChild(data);
            if (specs.length) {
                const script = document.createElement('script');
                script.textContent = `window.pptxFontsReady=Promise.all(JSON.parse(document.getElementById('pp-fonts').textContent).map(
f=>document.fonts.load(f.style+' '+f.weight+' 16px "'+f.alias+'"').then(faces=>{
if(!faces.length)throw new Error('Presentation font failed to load: '+f.family);
}))).then(()=>document.fonts.ready);`;
                document.head.appendChild(script);
            }
        }, { specs, css });
        await writeFile(htmlFile, await page.content());
    } finally {
        await page.close();
    }
}

export async function verifyFonts(page, session, slide, log, seen) {
    const { root } = await session.send('DOM.getDocument');
    const selector = `#slide-${slide + 1} [data-pp-font]`;
    const { nodeIds } = await session.send('DOM.querySelectorAll', { nodeId: root.nodeId, selector });
    const uses = await page.locator(selector).evaluateAll(elements => {
        const specs = JSON.parse(document.getElementById('pp-fonts').textContent);
        return elements.map(e => ({ ...specs[Number(e.dataset.ppFont)], text: e.textContent }));
    });
    for (let index = 0; index < nodeIds.length; index++) {
        const use = uses[index];
        const { fonts } = await session.send('CSS.getPlatformFontsForNode', { nodeId: nodeIds[index] });
        if (!fonts.length) continue;
        assert.ok(!/[a-z0-9]/i.test(use.text) || fonts.some(f => f.isCustomFont ||
            baseFamily(f.familyName).toLowerCase() === baseFamily(use.original).toLowerCase()),
            `Slide ${slide + 1}: unexpected font substitution for ${use.original}: ${fonts.map(f => f.familyName).join(', ')}`);
        const message = `${use.original} (${use.weight} ${use.style}) -> ${fonts.map(f => f.familyName + (f.isCustomFont ? ' [bundled]' : ' [installed]')).join(', ')}`;
        if (!seen.has(message)) { seen.add(message); log(`Font: ${message}`); }
    }
}
