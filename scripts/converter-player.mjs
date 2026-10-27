import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { build } from 'esbuild';

const require = createRequire(import.meta.url);
export const converterRoot = path.dirname(require.resolve('pptx2html/package.json'));

function correctSource(source, replacements) {
    for (const [before, after] of replacements) {
        assert.equal(source.split(before).length, 2, `The pinned PPTX2HTML source changed. Review the local animation correction for ${before.slice(0, 80)}.`);
        source = source.replace(before, () => after);
    }
    return source;
}

export function correctPlayerSource(source) {
    // The pinned upstream player clears authored run styles even for visibility-only animations.
    const replacements = [
        ['if (timing) this.collect();', `if (timing) this.collect();
    this.textStyles = new Map([...section.querySelectorAll('.pp-p span')].map((span) => [span, {
      color: span.style.color,
      webkitTextFillColor: span.style.webkitTextFillColor,
      fontWeight: span.style.fontWeight,
      fontStyle: span.style.fontStyle,
      textDecorationLine: span.style.textDecorationLine,
    }]));`],
        ["sp.style.color = st.color ? css(st.color) : '';",
            "sp.style.color = st.color ? css(st.color) : this.textStyles.get(sp).color;"],
        ["sp.style.webkitTextFillColor = st.color ? css(st.color) : '';",
            "sp.style.webkitTextFillColor = st.color ? css(st.color) : this.textStyles.get(sp).webkitTextFillColor;"],
        ["sp.style.fontWeight = st.bold === undefined ? '' : (st.bold ? 'bold' : 'normal');",
            "sp.style.fontWeight = st.bold === undefined ? this.textStyles.get(sp).fontWeight : (st.bold ? 'bold' : 'normal');"],
        ["sp.style.fontStyle = st.italic === undefined ? '' : (st.italic ? 'italic' : 'normal');",
            "sp.style.fontStyle = st.italic === undefined ? this.textStyles.get(sp).fontStyle : (st.italic ? 'italic' : 'normal');"],
        ["sp.style.textDecorationLine = st.underline ? 'underline' : '';",
            "sp.style.textDecorationLine = st.underline === undefined ? this.textStyles.get(sp).textDecorationLine : (st.underline ? 'underline' : 'none');"],
        [`if (sp.style.color || sp.style.fontWeight || sp.style.fontStyle || sp.style.textDecorationLine) {
          sp.style.color = ''; sp.style.webkitTextFillColor = ''; sp.style.fontWeight = ''; sp.style.fontStyle = ''; sp.style.textDecorationLine = '';
        }`, 'Object.assign(sp.style, this.textStyles.get(sp));']
    ];
    return correctSource(source, replacements);
}

export function correctTimingSource(source) {
    return correctSource(source, [
        [String.raw`/[MLCZEmlcze]|-?[0-9.]+(?:e[-+]?\d+)?/g`, String.raw`/[MLCZEmlcze]|-?[0-9.]+(?:[eE][-+]?\d+)?/g`],
        ['let d = dur / (info.spd || 1);', 'let d = dur / Math.abs(info.spd || 1);']
    ]);
}

export function correctEngineSource(source) {
    return correctSource(source, [
        ['const single = w.dur / (w.spd || 1);', 'const single = w.dur / Math.abs(w.spd || 1);'],
        ['p = ease(p, w.accel || 0, w.decel || 0);', 'p = ease(w.spd < 0 ? 1 - p : p, w.accel || 0, w.decel || 0);'],
        ['const single = b.dur / (b.spd || 1);', 'const single = b.dur / Math.abs(b.spd || 1);'],
        ['const single = b.dur === Infinity ? Infinity : b.dur / (b.spd || 1);', 'const single = b.dur === Infinity ? Infinity : b.dur / Math.abs(b.spd || 1);'],
        ['return { p: ease(p, b.accel || 0, b.decel || 0), running: false };', 'return { p: ease(b.spd < 0 ? 1 - p : p, b.accel || 0, b.decel || 0), running: false };'],
        ['return { p: ease(k / single, b.accel || 0, b.decel || 0), running: true };', 'return { p: ease(b.spd < 0 ? 1 - k / single : k / single, b.accel || 0, b.decel || 0), running: true };'],
        ['const per = (w.dur / (w.spd || 1)) * (w.autoRev ? 2 : 1);', 'const per = (w.dur / Math.abs(w.spd || 1)) * (w.autoRev ? 2 : 1);']
    ]);
}

export function correctStartupSource(source) {
    return correctSource(source, [[
        "window.PPTXPlayer = Player;",
        `window.PPTXPlayer = Player;
function start() {
  if (!window.pptxFontsReady) return autoStart();
  return window.pptxFontsReady.then(autoStart).catch(error => {
    document.body.textContent = 'This presentation could not load its fonts. Please reload the page.';
    console.error(error);
    throw error;
  });
}`
    ], ["document.addEventListener('DOMContentLoaded', autoStart);", "document.addEventListener('DOMContentLoaded', start);"],
    ['else autoStart();', 'else start();']]);
}

async function bundle(entry, browser = false, embeddedPlayer) {
    const result = await build({
        entryPoints: [path.join(converterRoot, ...entry)],
        bundle: true,
        minify: true,
        platform: browser ? 'browser' : 'node',
        format: browser ? 'iife' : 'esm',
        target: browser ? ['es2019'] : ['node24'],
        write: false,
        legalComments: 'none',
        plugins: [{
            name: 'correct-pinned-animations',
            setup(builder) {
                for (const [folder, file, correct] of [
                    ['runtime', 'player', correctPlayerSource], ['runtime', 'engine', correctEngineSource],
                    ['pptx', 'timing', correctTimingSource], ['runtime', 'standalone', correctStartupSource]
                ]) {
                    builder.onLoad({ filter: new RegExp(`[/\\\\]${folder}[/\\\\]${file}\\.js$`) }, async ({ path: file }) => ({
                        contents: correct(await readFile(file, 'utf8')), loader: 'js'
                    }));
                }
                if (embeddedPlayer) builder.onLoad({ filter: /[/\\]generated[/\\]player-source\.js$/ }, () => ({
                    contents: `export const PLAYER_SOURCE = ${JSON.stringify(embeddedPlayer)};`, loader: 'js'
                }));
            }
        }]
    });
    return result.outputFiles[0].text.trim();
}

let player;
export function correctedPlayer() {
    player ??= bundle(['src', 'runtime', 'standalone.js'], true);
    return player;
}

let converter;
export function correctedConverter() {
    converter ??= correctedPlayer().then(player => bundle(['bin', 'pptx2html.js'], false, player));
    return converter;
}

let engine;
export function correctedEngine() {
    engine ??= bundle(['src', 'runtime', 'engine.js']).then(source => import(`data:text/javascript;base64,${Buffer.from(source).toString('base64')}`));
    return engine;
}
