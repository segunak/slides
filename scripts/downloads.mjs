import { execFile } from 'node:child_process';
import { stat } from 'node:fs/promises';
import assert from 'node:assert/strict';
import { promisify } from 'node:util';

const exec = promisify(execFile);
const repository = 'segunak/slides';
const releaseTag = 'downloads';
const releaseUrl = `https://github.com/${repository}/releases/download/${releaseTag}/`;
export const downloadMarker = 'data-slide-downloads';

export const releaseAssetUrl = name => releaseUrl + encodeURIComponent(name);
export const mediaAssetPrefix = name => `${name}--media--`;

export function releasedMediaNames(html, name) {
    const urls = html.match(/https:\/\/github\.com\/segunak\/slides\/releases\/download\/downloads\/[^"'<>\\\s)]+/g) || [];
    return [...new Set(urls.filter(url => url.startsWith(releaseAssetUrl(mediaAssetPrefix(name))))
        .map(url => decodeURIComponent(url.slice(releaseUrl.length))))];
}

export function validateDownloadSize(bytes, name) {
    assert.ok(bytes > 0 && bytes < 2 ** 31, `Download ${name} must be nonempty and smaller than GitHub's 2 GiB release-asset limit.`);
}

export function downloadLinks(name) {
    const base = releaseAssetUrl(name);
    return `<a href="${base}.pdf" aria-label="Download PDF">Download PDF</a> <a href="${base}.pptx" aria-label="Download PowerPoint">Download PowerPoint</a>`;
}

export function addDownloadControls(html, name) {
    if (html.includes(downloadMarker)) throw new Error('The presentation already contains download controls.');
    const controls = `<style>
.pp-ui a{background:rgba(40,40,40,.9);color:#fff;border-radius:4px;padding:7px 9px;text-decoration:none;white-space:nowrap}
.pp-ui a:hover,.pp-ui a:focus-visible{background:rgba(80,80,80,.95)}
.pp-ui{flex-wrap:wrap;max-width:calc(100vw - 24px)}
</style>
<script ${downloadMarker}>
function addSlideDownloads(){
  const ui=document.querySelector('.pp-ui');
  if(!ui)throw new Error('Presentation download toolbar is unavailable.');
  ui.insertAdjacentHTML('beforeend',${JSON.stringify(downloadLinks(name))});
}
function ready(){if(window.pptxFontsReady)window.pptxFontsReady.then(addSlideDownloads);else addSlideDownloads();}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',ready,{once:true});
else ready();
</script>`;
    if (!html.includes('</body>')) throw new Error('The converted presentation has no body.');
    return html.replace('</body>', controls + '\n</body>');
}

async function gh(args) {
    const result = await exec('gh', args, { encoding: 'utf8', timeout: 300_000, maxBuffer: 2 * 1024 * 1024, windowsHide: true });
    return result.stdout;
}

export function releaseDownloads(run = gh, log = console.log) {
    let release;
    return {
        async list() {
            try {
                release = JSON.parse(await run(['api', `repos/${repository}/releases/tags/${releaseTag}`]));
            } catch (error) {
                if (!/\(HTTP 404\)/.test(error.stderr || '')) {
                    throw new Error('Could not read the Slide Downloads release. Check the GitHub App Contents permission.');
                }
                release = null;
                return new Set();
            }
            if (release.immutable) throw new Error('The downloads release is immutable. Rolling downloads require a mutable release.');
            const pages = JSON.parse(await run(['api', `repos/${repository}/releases/${release.id}/assets`, '--paginate', '--slurp']));
            return new Set(pages.flat().map(asset => asset.name));
        },
        async publish(files, removed) {
            for (const file of files) {
                for (const source of [file.pptx, file.pdf, ...(file.media || [])]) {
                    const info = await stat(source);
                    assert.ok(info.isFile(), `Download is not a regular file: ${source}.`);
                    validateDownloadSize(info.size, file.name);
                }
            }
            if (!release && !files.length) return;
            if (!release) {
                await run(['release', 'create', releaseTag, '--repo', repository, '--target', 'main',
                    '--title', 'Slide Downloads', '--notes', 'Latest editable PowerPoints and printable PDFs. View slides at https://segunakinyemi.com/slides/', '--latest']);
                release = { id: null };
            }
            for (const file of files) {
                for (let attempt = 1; ; attempt++) {
                    try {
                        await run(['release', 'upload', releaseTag, file.pptx, file.pdf, ...(file.media || []),
                            '--clobber', '--repo', repository]);
                        break;
                    } catch {
                        if (attempt === 3) throw new Error(`Could not replace downloads for ${file.name} after three attempts. No website publication will be committed.`);
                        log(`Retrying download upload for ${file.name} (${attempt + 1}/3).`);
                    }
                }
            }
            for (const name of removed) {
                await run(['release', 'delete-asset', releaseTag, name, '--yes', '--repo', repository]);
            }
        }
    };
}
