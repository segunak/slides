import assert from 'node:assert/strict';
import { mkdtemp, readFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { test } from 'node:test';
import { connectOneDrive } from './onedrive.mjs';

const mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const item = { id: 'drive!item', name: 'Deck.pptx', size: 4, eTag: '"v1"', file: { mimeType } };

function fixture() {
    const calls = [];
    const logs = [];
    const cache = {
        deserialize: value => calls.push(['restore', value]),
        serialize: () => 'renewed-cache',
        getAllAccounts: async () => [{ homeAccountId: 'account' }]
    };
    const client = {
        getTokenCache: () => cache,
        acquireTokenSilent: async request => {
            calls.push(['silent', request]);
            return { accessToken: 'test-access-token' };
        }
    };
    const options = {
        env: { ONEDRIVE_CLIENT_ID: 'client-id', ONEDRIVE_AUTH_CACHE: 'saved-cache', GH_TOKEN: 'app-token' },
        createClient: config => {
            calls.push(['client', config]);
            return client;
        },
        saveCache: async value => calls.push(['save', value]),
        request: async (url, options) => {
            calls.push(['request', url, options]);
            return Response.json(item);
        },
        log: message => logs.push(message)
    };
    return { options, calls, logs, cache, client };
}

test('uses native personal-account MSAL and persists renewal before any Graph request', async () => {
    const f = fixture();
    const drive = await connectOneDrive(f.options);
    assert.deepEqual(f.calls.find(call => call[0] === 'save'), ['save', 'renewed-cache']);
    assert.ok(!f.calls.some(call => call[0] === 'request'));
    const config = f.calls.find(call => call[0] === 'client')[1];
    assert.equal(config.auth.authority, 'https://login.microsoftonline.com/consumers');
    assert.equal(config.auth.clientSecret, undefined);
    assert.deepEqual(f.calls.find(call => call[0] === 'silent')[1].scopes, ['Files.Read']);
    assert.equal(await drive.getItem(item.id).then(result => result.id), item.id);
    const [, url, options] = f.calls.find(call => call[0] === 'request');
    assert.equal(url.origin, 'https://graph.microsoft.com');
    assert.ok(url.pathname.endsWith('drive!item'));
    assert.equal(options.redirect, 'manual');
    assert.equal(options.headers.Authorization, 'Bearer test-access-token');
    assert.doesNotMatch(f.logs.join('\n'), /saved-cache|renewed-cache|test-access-token|app-token/);
});

test('does not rewrite an unchanged authentication cache', async () => {
    const f = fixture();
    f.cache.serialize = () => 'saved-cache';
    await connectOneDrive(f.options);
    assert.ok(!f.calls.some(call => call[0] === 'save'));
});

for (const name of ['ONEDRIVE_CLIENT_ID', 'ONEDRIVE_AUTH_CACHE', 'GH_TOKEN']) {
    test(`reports missing ${name} before authenticating`, async () => {
        const f = fixture();
        delete f.options.env[name];
        await assert.rejects(connectOneDrive(f.options), new RegExp(name));
        assert.deepEqual(f.calls, []);
    });
}

for (const failure of ['cache', 'accounts', 'token', 'missing-token']) {
    test(`reports ${failure} failure without exposing credentials`, async () => {
        const f = fixture();
        if (failure === 'cache') f.cache.deserialize = () => { throw new Error('saved-cache'); };
        if (failure === 'accounts') f.cache.getAllAccounts = async () => [];
        if (failure === 'token') f.client.acquireTokenSilent = async () => { throw new Error('test-access-token'); };
        if (failure === 'missing-token') f.client.acquireTokenSilent = async () => null;
        await assert.rejects(connectOneDrive(f.options), error => {
            assert.match(error.message, /Reconnect Personal Slide Publisher/);
            assert.doesNotMatch(error.message, /saved-cache|test-access-token/);
            return true;
        });
        assert.ok(!f.calls.some(call => call[0] === 'save' || call[0] === 'request'));
    });
}

test('secret persistence failure stops the publisher before downloads', async () => {
    const f = fixture();
    f.options.saveCache = async () => { throw new Error('Secret write refused'); };
    await assert.rejects(connectOneDrive(f.options), /Secret write refused/);
    assert.ok(!f.calls.some(call => call[0] === 'request'));
});

for (const status of [401, 403, 404, 429, 500]) {
    test(`metadata HTTP ${status} fails explicitly after authentication renewal`, async () => {
        const f = fixture();
        f.options.request = async () => new Response(null, { status });
        const drive = await connectOneDrive(f.options);
        await assert.rejects(drive.getItem(item.id), new RegExp(`HTTP ${status}`));
        assert.ok(f.calls.some(call => call[0] === 'save'));
    });
}

for (const invalid of [
    { ...item, id: 'wrong' }, { ...item, name: 'Deck.pdf' }, { ...item, file: {} },
    { ...item, size: -1 }, { ...item, size: '4' }, { ...item, eTag: '' }
]) {
    test(`rejects invalid metadata ${JSON.stringify(invalid)}`, async () => {
        const f = fixture();
        f.options.request = async () => Response.json(invalid);
        const drive = await connectOneDrive(f.options);
        await assert.rejects(drive.getItem(item.id), /invalid PowerPoint metadata/);
    });
}

test('network failures do not expose an authorization value', async () => {
    const f = fixture();
    f.options.request = async () => { throw new Error('test-access-token'); };
    const drive = await connectOneDrive(f.options);
    await assert.rejects(drive.getItem(item.id), error => {
        assert.match(error.message, /request failed or timed out/);
        assert.doesNotMatch(error.message, /test-access-token/);
        return true;
    });
});

test('downloads to a temporary file without forwarding the bearer token to the signed URL', async t => {
    const directory = await mkdtemp(path.join(tmpdir(), 'slides-download-test-'));
    t.after(() => rm(directory, { recursive: true, force: true }));
    const f = fixture();
    const requests = [];
    f.options.request = async (url, options) => {
        requests.push([url, options]);
        if (url.pathname.endsWith('/content')) {
            return new Response(null, { status: 302, headers: { location: 'https://download.example/deck?signature=private' } });
        }
        if (url.hostname === 'download.example') return new Response('pptx');
        return Response.json(item);
    };
    const drive = await connectOneDrive(f.options);
    const destination = path.join(directory, 'source.pptx');
    await drive.download(item, destination);
    assert.equal(await readFile(destination, 'utf8'), 'pptx');
    assert.equal(requests[0][1].headers.Authorization, 'Bearer test-access-token');
    assert.equal(requests[1][1].headers, undefined);
    assert.equal(requests[2][0].origin, 'https://graph.microsoft.com');
    assert.deepEqual(await readdir(directory), ['source.pptx']);
    assert.doesNotMatch(f.logs.join('\n'), /signature|private/);
});

for (const mode of ['unsafe-redirect', 'http-error', 'short-download', 'edited-during-download', 'download-error']) {
    test(`rejects ${mode} without leaking signed URLs`, async t => {
        const directory = await mkdtemp(path.join(tmpdir(), 'slides-download-test-'));
        t.after(() => rm(directory, { recursive: true, force: true }));
        const f = fixture();
        f.options.request = async url => {
            if (url.pathname.endsWith('/content')) {
                if (mode === 'http-error') return new Response(null, { status: 403 });
                return new Response(null, { status: 302, headers: {
                    location: `${mode === 'unsafe-redirect' ? 'http' : 'https'}://download.example/deck?signature=private`
                } });
            }
            if (url.hostname === 'download.example') {
                if (mode === 'download-error') throw new Error('https://download.example/?signature=private');
                return new Response(mode === 'short-download' ? 'bad' : 'pptx');
            }
            return Response.json(mode === 'edited-during-download' ? { ...item, eTag: '"v2"' } : item);
        };
        const drive = await connectOneDrive(f.options);
        await assert.rejects(drive.download(item, path.join(directory, 'source.pptx')), error => {
            assert.match(error.message, /invalid download redirect|HTTP 403|incomplete|changed during download|download failed/);
            assert.doesNotMatch(error.message, /signature=private/);
            return true;
        });
    });
}
