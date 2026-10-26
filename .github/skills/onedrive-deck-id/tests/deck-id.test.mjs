import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { access, readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';
import { parseDocument } from 'yaml';
import { connectLookup, lookupDecks, resolvePublicLink, runCli } from '../scripts/deck-id.mjs';
import { createDriveReader, drivePath } from '../../../../scripts/onedrive.mjs';

const clientId = '6a943419-c826-401b-b87a-24fa32dc60fe';
const fileId = '750D396C5CADCEBD!s67f4c2a13c6d48f1b533ba6f8d574bd0';
const share = 'https://1drv.ms/p/c/750d396c5cadcebd/IQShwvRnbTzxSLUzum-NV0vQAartnDl4cQJ42eL2ccVtiwk?em=2&wdAr=1.7777777777777777';
const expanded = `https://onedrive.live.com/edit?cid=750d396c5cadcebd&id=${encodeURIComponent(fileId)}&resid=${encodeURIComponent(fileId)}`;
const item = {
    id: fileId, name: 'How to Train Your AI.pptx', eTag: '"v1"', size: 137401881,
    file: { mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }
};

test('expands the existing embed format without fetching the final viewer or decoding the sharing token', async () => {
    const calls = [];
    const id = await resolvePublicLink(share, async (url, options) => {
        calls.push([url, options]);
        return new Response(null, { status: 301, headers: { location: expanded } });
    });
    assert.equal(id, fileId);
    assert.equal(calls.length, 1);
    assert.equal(calls[0][0].href, share);
    assert.equal(calls[0][1].redirect, 'manual');
    assert.equal(calls[0][1].headers, undefined);
    assert.ok(calls[0][1].signal instanceof AbortSignal);
});

for (const parameter of ['id', 'resid']) {
    test(`reads an expanded ${parameter} exactly once with no network request`, async () => {
        const url = `https://onedrive.live.com/embed?${parameter}=${encodeURIComponent('opaque!id+%part/with?punctuation')}`;
        assert.equal(await resolvePublicLink(url, () => assert.fail('No request needed')), 'opaque!id+%part/with?punctuation');
    });
}

test('accepts pasted iframe markup and unescapes its URL', async () => {
    const iframe = `<iframe data-speaking-talk-slides-src="${share.replaceAll('&', '&amp;')}" title="Deck"></iframe>`;
    assert.equal(await resolvePublicLink(iframe, async url => {
        assert.equal(url.href, share);
        return new Response(null, { status: 302, headers: { location: expanded } });
    }), fileId);
});

test('follows relative redirects on allowed hosts', async () => {
    let calls = 0;
    assert.equal(await resolvePublicLink('https://1drv.ms/p/short', async url => {
        calls++;
        if (calls === 1) {
            assert.equal(url.pathname, '/p/short');
            return new Response(null, { status: 307, headers: { location: '/p/next' } });
        }
        assert.equal(url.pathname, '/p/next');
        return new Response(null, { status: 302, headers: { location: expanded } });
    }), fileId);
    assert.equal(calls, 2);
});

for (const value of [
    'not-a-link', 'http://1drv.ms/p/short', 'https://evil.example/file',
    'https://1drv.ms.evil.example/file', 'https://user:password@1drv.ms/p/short',
    'https://1drv.ms:444/p/short', 'file:///C:/private.pptx'
]) {
    test(`rejects unsupported link ${value} without making a request`, async () => {
        await assert.rejects(resolvePublicLink(value, () => assert.fail('No request allowed')), /Expected a public|Only HTTPS/);
    });
}

test('rejects a redirect outside personal OneDrive before requesting it', async () => {
    let calls = 0;
    await assert.rejects(resolvePublicLink(share, async () => {
        calls++;
        return new Response(null, { status: 302, headers: { location: 'https://evil.example/?id=guessed' } });
    }), /Only HTTPS/);
    assert.equal(calls, 1);
});

test('never mistakes cid or the shortened sharing token for a file ID', async () => {
    await assert.rejects(resolvePublicLink('https://onedrive.live.com/?cid=account', async () => new Response('viewer')), /did not expose a file ID/);
});

test('rejects conflicting file IDs rather than picking one', async () => {
    await assert.rejects(resolvePublicLink('https://onedrive.live.com/?id=first&resid=second'), /conflicting file IDs/);
});

test('bounds redirects and reports network or non-redirect failures explicitly', async () => {
    let calls = 0;
    await assert.rejects(resolvePublicLink(share, async () => {
        calls++;
        return new Response(null, { status: 302, headers: { location: share } });
    }), /eight redirects/);
    assert.equal(calls, 8);
    await assert.rejects(resolvePublicLink(share, async () => { throw new Error('private network detail'); }), /Could not expand/);
    await assert.rejects(resolvePublicLink(share, async () => new Response(null, { status: 403 })), /no ID was guessed/);
    await assert.rejects(resolvePublicLink(share, async () => new Response(null, { status: 302 })), /did not expose/);
});

test('exact OneDrive paths encode reserved characters and Windows separators without guessing', async () => {
    const requests = [];
    const drive = createDriveReader('read-only-token', async (url, options) => {
        requests.push([url, options]);
        return Response.json(item);
    });
    const filePath = 'Documents\\Talks\\Résumé #1 + 100%.pptx';
    const result = await drive.getItemByPath(filePath);
    assert.equal(result.id, fileId);
    assert.equal(requests[0][0].pathname, '/v1.0/me/drive/root:/Documents/Talks/R%C3%A9sum%C3%A9%20%231%20%2B%20100%25.pptx');
    assert.equal(requests[0][0].searchParams.get('$select'), 'id,name,eTag,size,file,webUrl');
    assert.equal(requests[0][1].headers.Authorization, 'Bearer read-only-token');
});

for (const invalid of ['', '../Deck.pptx', 'Folder/../Deck.pptx', 'Folder//Deck.pptx', 'C:\\Users\\Deck.pptx', 'Folder/./Deck.pptx']) {
    test(`rejects invalid OneDrive-relative path ${invalid}`, () => assert.throws(() => drivePath(invalid), /relative to your OneDrive/));
}

function authentication() {
    const calls = [];
    const reports = [];
    const client = {
        acquireTokenByDeviceCode: async options => {
            calls.push(['device', options]);
            options.deviceCodeCallback({ message: 'Visit Microsoft and enter EXAMPLE-CODE.' });
            return { accessToken: 'test-access-token' };
        },
        getTokenCache: () => ({
            deserialize: value => calls.push(['restore', value]),
            getAllAccounts: async () => [{ homeAccountId: 'test-account' }],
            serialize: () => assert.fail('Lookup must not persist tokens')
        }),
        acquireTokenSilent: async options => {
            calls.push(['silent', options]);
            return { accessToken: 'test-access-token' };
        }
    };
    const options = {
        env: { ONEDRIVE_CLIENT_ID: clientId },
        clientIdFromGitHub: () => { calls.push(['variable']); return clientId; },
        createClient: config => { calls.push(['client', config]); return client; },
        request: async (url, options) => { calls.push(['graph', url, options]); return Response.json(item); },
        report: message => reports.push(message)
    };
    return { client, calls, reports, options };
}

test('local lookup signs in once with native read-only MSAL and keeps credentials in memory', async () => {
    const f = authentication();
    const drive = await connectLookup(f.options);
    assert.equal((await drive.getItem(fileId)).id, fileId);
    const config = f.calls.find(call => call[0] === 'client')[1];
    assert.equal(config.auth.authority, 'https://login.microsoftonline.com/consumers');
    assert.equal(config.auth.clientSecret, undefined);
    assert.deepEqual(f.calls.find(call => call[0] === 'device')[1].scopes, ['Files.Read']);
    assert.ok(f.reports.some(message => message.includes('EXAMPLE-CODE')));
    assert.doesNotMatch(f.reports.join('\n'), /test-access-token/);
    assert.ok(!f.calls.some(call => call[0] === 'variable' || call[0] === 'silent'));
});

test('can read the public client ID variable without requiring a GitHub App token', async () => {
    const f = authentication();
    f.options.env = {};
    await connectLookup(f.options);
    assert.equal(f.calls.filter(call => call[0] === 'variable').length, 1);
});

test('uses an explicitly supplied in-memory authentication cache without changing it', async () => {
    const f = authentication();
    f.options.env.ONEDRIVE_AUTH_CACHE = 'test-cache';
    f.options.env.GITHUB_ACTIONS = 'true';
    await connectLookup(f.options);
    assert.ok(f.calls.some(call => call[0] === 'silent'));
    assert.ok(!f.calls.some(call => call[0] === 'device'));
    assert.deepEqual(f.reports, []);
});

test('does not attempt an interactive fallback when a supplied cache fails', async () => {
    const f = authentication();
    f.options.env.ONEDRIVE_AUTH_CACHE = 'test-cache';
    f.client.acquireTokenSilent = async () => { throw new Error('private token'); };
    await assert.rejects(connectLookup(f.options), error => {
        assert.match(error.message, /could not sign in silently/);
        assert.doesNotMatch(error.message, /private token/);
        return true;
    });
    assert.ok(!f.calls.some(call => call[0] === 'device'));
});

test('CI requires provided authentication instead of hanging on a sign-in prompt', async () => {
    const f = authentication();
    f.options.env.CI = 'true';
    await assert.rejects(connectLookup(f.options), /disabled in CI/);
    assert.ok(!f.calls.some(call => call[0] === 'device'));
});

test('invalid app ID and failed sign-in produce clear errors without raw credential responses', async () => {
    const f = authentication();
    f.options.env.ONEDRIVE_CLIENT_ID = 'not-an-id';
    await assert.rejects(connectLookup(f.options), /Application \(client\) ID/);
    f.options.env.ONEDRIVE_CLIENT_ID = clientId;
    f.client.acquireTokenByDeviceCode = async () => { throw new Error('private response'); };
    await assert.rejects(connectLookup(f.options), /sign-in failed or expired/);
    f.client.acquireTokenByDeviceCode = async () => null;
    await assert.rejects(connectLookup(f.options), /No IDs were verified/);
});

test('returns only canonical metadata, deduplicating multiple inputs for the same deck', async () => {
    const requests = [];
    const drive = {
        getItemByPath: async filePath => { requests.push(['path', filePath]); return item; },
        getItem: async id => { requests.push(['id', id]); return item; }
    };
    const results = await lookupDecks([{ path: 'Documents/Deck.pptx' }, { id: fileId }], drive);
    assert.deepEqual(results, [{ id: fileId, name: item.name, size: item.size }]);
    assert.deepEqual(requests, [['path', 'Documents/Deck.pptx'], ['id', fileId], ['id', fileId], ['id', fileId]]);
});

test('requires the second canonical lookup to resolve to the same file version', async () => {
    for (const difference of [{ id: 'different-id' }, { name: 'Other.pptx' }, { eTag: '"v2"' }, { size: 5 }]) {
        let calls = 0;
        await assert.rejects(lookupDecks([{ id: fileId }], {
            getItem: async () => ++calls === 1 ? item : { ...item, ...difference }
        }), /changed during verification/);
    }
});

test('link and exact-path CLI modes return parseable verified metadata using one connection', async () => {
    const requests = [];
    const output = [];
    let connections = 0;
    await runCli([share, '--path', 'Documents\\Deck.pptx'], {
        resolve: async input => { assert.equal(input, share); return fileId; },
        connect: async () => {
            connections++;
            return createDriveReader('read-only-token', async url => {
                requests.push(url);
                return Response.json(item);
            });
        },
        write: value => output.push(value)
    });
    assert.equal(connections, 1);
    assert.deepEqual(JSON.parse(output.join('')), [{ id: fileId, name: item.name, size: item.size }]);
    assert.equal(requests.filter(url => url.pathname.includes('/items/')).length, 3);
    assert.equal(requests.filter(url => url.pathname.includes('/root:')).length, 1);
});

test('CLI refuses invalid inputs before authenticating, and never prints partial success', async () => {
    const options = { connect: () => assert.fail('No authentication expected'), write: () => assert.fail('No output expected') };
    await assert.rejects(runCli([], options), /Provide a public/);
    await assert.rejects(runCli(['--unknown'], options), /Unknown option/);
    await assert.rejects(runCli(['--path', 'C:\\Deck.pptx'], options), /relative to your OneDrive/);
    await assert.rejects(runCli(['https://invalid.example/'], options), /Only HTTPS/);
    await assert.rejects(runCli([share], {
        resolve: async () => fileId,
        connect: async () => ({ getItem: async () => { throw new Error('Graph HTTP 404'); } }),
        write: options.write
    }), /Graph HTTP 404/);
});

test('the actual command has self-contained help and a nonzero exit for missing input', () => {
    const script = fileURLToPath(new URL('../scripts/deck-id.mjs', import.meta.url));
    const output = execFileSync(process.execPath, [script, '--help'], { encoding: 'utf8' });
    assert.match(output, /npm run deck-id/);
    assert.match(output, /--path/);
    assert.throws(() => execFileSync(process.execPath, [script], { stdio: 'pipe' }), error => {
        assert.equal(error.status, 1);
        assert.match(error.stderr.toString(), /Provide a public OneDrive link/);
        assert.equal(error.stdout.toString(), '');
        return true;
    });
});

test('skill metadata, resource links, and npm entry points remain discoverable after migration', async () => {
    const skillUrl = new URL('../SKILL.md', import.meta.url);
    const skill = await readFile(skillUrl, 'utf8');
    const frontmatter = skill.match(/^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/);
    assert.ok(frontmatter, 'Missing skill frontmatter.');
    const doc = parseDocument(frontmatter[1]);
    assert.deepEqual(doc.errors, []);
    const metadata = doc.toJS();
    assert.equal(metadata.name, 'onedrive-deck-id');
    assert.match(metadata.name, /^[a-z0-9]+(?:-[a-z0-9]+)*$/);
    assert.ok(metadata.name.length <= 64);
    assert.ok(typeof metadata.description === 'string' && metadata.description.length > 0 && metadata.description.length <= 1024);
    assert.match(metadata.description, /OneDrive/);
    assert.ok(skill.split(/\r?\n/).length < 500);
    for (const [, link] of skill.matchAll(/\[[^\]]+\]\(([^)]+)\)/g)) {
        if (!/^https?:/.test(link)) await access(new URL(link, skillUrl));
    }
    const pkg = JSON.parse(await readFile(new URL('../../../../package.json', import.meta.url), 'utf8'));
    assert.equal(pkg.scripts['deck-id'], 'node .github/skills/onedrive-deck-id/scripts/deck-id.mjs');
    assert.ok(pkg.scripts.test.includes('.github/skills/onedrive-deck-id/tests/*.test.mjs'));
    for (const oldPath of ['scripts/deck-id.mjs', 'scripts/deck-id.test.mjs']) {
        await assert.rejects(access(new URL(`../../../../${oldPath}`, import.meta.url)), { code: 'ENOENT' });
    }
});
