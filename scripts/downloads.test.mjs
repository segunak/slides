import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { test } from 'node:test';
import { addDownloadControls, downloadLinks, releaseDownloads, validateDownloadSize,
    releaseAssetUrl, mediaAssetPrefix, releasedMediaNames } from './downloads.mjs';

async function pair(t, name = 'alpha') {
    const root = await mkdtemp(path.join(tmpdir(), 'slides-download-test-'));
    t.after(() => rm(root, { recursive: true, force: true }));
    const files = { name, pptx: path.join(root, `${name}.pptx`), pdf: path.join(root, `${name}.pdf`) };
    await writeFile(files.pptx, 'original PowerPoint bytes');
    await writeFile(files.pdf, '%PDF-fixture');
    return files;
}

function fixture(release = { id: 42 }, assets = ['alpha.pptx', 'alpha.pdf']) {
    const calls = [];
    const messages = [];
    let uploadFailures = 0;
    const downloads = releaseDownloads(async args => {
        calls.push(args);
        if (args[0] === 'api' && args[1].endsWith('/tags/downloads')) {
            if (release === null) throw Object.assign(new Error('not found'), { stderr: 'gh: Not Found (HTTP 404)' });
            return JSON.stringify(release);
        }
        if (args[0] === 'api') return JSON.stringify([assets.map(name => ({ name }))]);
        if (args[1] === 'upload' && uploadFailures-- > 0) throw new Error('upload interrupted');
        return '';
    }, message => messages.push(message));
    return { downloads, calls, messages, failUploads: count => { uploadFailures = count; } };
}

test('uses stable direct downloads in both formats and injects controls without replacing the player', () => {
    const links = downloadLinks('alpha');
    assert.match(links, /https:\/\/github\.com\/segunak\/slides\/releases\/download\/downloads\/alpha\.pdf/);
    assert.match(links, /alpha\.pptx/);
    const page = addDownloadControls('<body><script>originalPlayer()</script></body>', 'alpha');
    assert.match(page, /originalPlayer/);
    assert.match(page, /data-slide-downloads/);
    assert.match(page, /DOMContentLoaded/);
    assert.match(page, /document\.querySelector\('\.pp-ui'\)/);
    assert.throws(() => addDownloadControls(page, 'alpha'), /already contains/);
    assert.throws(() => addDownloadControls('not a document', 'alpha'), /no body/);
});

test('updates one rolling release in place without creating new releases or accumulating asset versions', async t => {
    const f = fixture();
    assert.deepEqual([...await f.downloads.list()], ['alpha.pptx', 'alpha.pdf']);
    const files = [await pair(t)];
    await f.downloads.publish(files, []);
    await f.downloads.publish(files, []);
    assert.ok(!f.calls.some(args => args[1] === 'create'));
    const uploads = f.calls.filter(args => args[1] === 'upload');
    assert.equal(uploads.length, 2);
    assert.deepEqual(uploads[0], ['release', 'upload', 'downloads', files[0].pptx, files[0].pdf, '--clobber', '--repo', 'segunak/slides']);
    assert.deepEqual(uploads[0], uploads[1]);
});

test('creates a single downloads release only when publishing initial files', async t => {
    const f = fixture(null);
    assert.equal((await f.downloads.list()).size, 0);
    await f.downloads.publish([], []);
    assert.ok(!f.calls.some(args => args[1] === 'create'));
    await f.downloads.publish([await pair(t)], []);
    await f.downloads.publish([await pair(t, 'beta')], []);
    assert.equal(f.calls.filter(args => args[1] === 'create').length, 1);
});

test('upload retries are bounded and failed replacement never reports success', async t => {
    const f = fixture();
    await f.downloads.list();
    f.failUploads(2);
    await f.downloads.publish([await pair(t)], []);
    assert.equal(f.calls.filter(args => args[1] === 'upload').length, 3);
    f.failUploads(3);
    await assert.rejects(f.downloads.publish([await pair(t, 'beta')], []), /three attempts/);
});

test('validates every download before creating or replacing release assets', async t => {
    assert.throws(() => validateDownloadSize(0, 'empty'), /nonempty/);
    assert.throws(() => validateDownloadSize(2 ** 31, 'too-large'), /2 GiB/);
    assert.throws(() => validateDownloadSize(2 ** 31 + 1, 'too-large'), /2 GiB/);
    assert.doesNotThrow(() => validateDownloadSize(2 ** 31 - 1, 'allowed'));
    const f = fixture(null);
    await f.downloads.list();
    const valid = await pair(t);
    const empty = await pair(t, 'empty');
    await writeFile(empty.pdf, '');
    await assert.rejects(f.downloads.publish([valid, empty], []), /nonempty/);
    assert.ok(!f.calls.some(args => args[0] === 'release'));
});

test('removes only the named assets for deliberately unpublished decks', async () => {
    const f = fixture();
    await f.downloads.list();
    await f.downloads.publish([], ['alpha.pptx', 'alpha.pdf']);
    assert.deepEqual(f.calls.filter(args => args[1] === 'delete-asset'), [
        ['release', 'delete-asset', 'downloads', 'alpha.pptx', '--yes', '--repo', 'segunak/slides'],
        ['release', 'delete-asset', 'downloads', 'alpha.pdf', '--yes', '--repo', 'segunak/slides']
    ]);
});

test('immutable releases and read failures are explicit errors, not empty successful listings', async () => {
    await assert.rejects(fixture({ id: 42, immutable: true }).downloads.list(), /immutable/);
    const downloads = releaseDownloads(async () => { throw new Error('permission denied'); });
    await assert.rejects(downloads.list(), /Could not read/);
});

test('uploads stable media with its deck and applies the same validation, replacement and retries', async t => {
    const f = fixture();
    await f.downloads.list();
    const file = await pair(t);
    file.media = [path.join(path.dirname(file.pptx), 'alpha--media--media1.mp4')];
    await writeFile(file.media[0], 'unaltered media');
    f.failUploads(1);
    await f.downloads.publish([file], ['alpha--media--old.mp4']);
    const uploads = f.calls.filter(args => args[1] === 'upload');
    assert.equal(uploads.length, 2);
    assert.deepEqual(uploads[0], ['release', 'upload', 'downloads', file.pptx, file.pdf, ...file.media,
        '--clobber', '--repo', 'segunak/slides']);
    assert.deepEqual(uploads[0], uploads[1]);
    assert.ok(f.calls.findIndex(args => args[1] === 'delete-asset') > f.calls.findLastIndex(args => args[1] === 'upload'));
    const count = uploads.length;
    await writeFile(file.media[0], '');
    await assert.rejects(f.downloads.publish([file], []), /nonempty/);
    assert.equal(f.calls.filter(args => args[1] === 'upload').length, count);
});

test('reads owned release-media references without another manifest or confusion between similar deck names', () => {
    const url = releaseAssetUrl('alpha--media--media1.mp4');
    assert.equal(mediaAssetPrefix('alpha'), 'alpha--media--');
    assert.deepEqual(releasedMediaNames(`<video src="${url}"></video>
<script>{"url":"${url}"}</script>
<img src="${releaseAssetUrl('alpha-beta--media--image.png')}">
<a href="${releaseAssetUrl('alpha.pptx')}">Download</a>`, 'alpha'), ['alpha--media--media1.mp4']);
    assert.deepEqual(releasedMediaNames('<video src="assets/media1.mp4">', 'alpha'), []);
});
