import { execFileSync } from 'node:child_process';
import { createWriteStream } from 'node:fs';
import { stat } from 'node:fs/promises';
import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { PublicClientApplication } from '@azure/msal-node';

const repository = 'segunak/slides';
const mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

function required(env, name) {
    if (!env[name]?.trim()) throw new Error(`Missing ${name}. Configure the slides repository's Actions settings.`);
    return env[name];
}

export function oneDriveClientConfig(clientId) {
    return {
        auth: { clientId, authority: 'https://login.microsoftonline.com/consumers' },
        system: { loggerOptions: { piiLoggingEnabled: false, loggerCallback: () => {} } }
    };
}

export function drivePath(filePath) {
    const parts = filePath.replaceAll('\\', '/').replace(/^\/+/, '').split('/');
    if (parts.some(part => !part || part === '.' || part === '..' || /[:\u0000-\u001f]/.test(part))) {
        throw new Error('Use the file path relative to your OneDrive root, not a local disk path or a path containing "." or "..".');
    }
    return `/root:/${parts.map(encodeURIComponent).join('/')}`;
}

export function saveAuthCache(cache) {
    try {
        execFileSync('gh', ['secret', 'set', 'ONEDRIVE_AUTH_CACHE', '--repo', repository], {
            input: cache,
            encoding: 'utf8',
            stdio: ['pipe', 'pipe', 'pipe'],
            timeout: 30_000,
            windowsHide: true
        });
    } catch {
        throw new Error('Could not save renewed OneDrive credentials. The GitHub App needs Secrets read/write access to segunak/slides.');
    }
}

export async function connectOneDrive({
    env = process.env,
    createClient = config => new PublicClientApplication(config),
    saveCache = saveAuthCache,
    request = fetch,
    log = console.log
} = {}) {
    const clientId = required(env, 'ONEDRIVE_CLIENT_ID');
    const savedCache = required(env, 'ONEDRIVE_AUTH_CACHE');
    required(env, 'GH_TOKEN');
    const client = createClient(oneDriveClientConfig(clientId));
    let auth;
    try {
        client.getTokenCache().deserialize(savedCache);
        const accounts = await client.getTokenCache().getAllAccounts();
        if (accounts.length !== 1) throw new Error('Expected one connected Microsoft account.');
        auth = await client.acquireTokenSilent({ account: accounts[0], scopes: ['Files.Read'] });
        if (!auth?.accessToken) throw new Error('Missing Microsoft access token.');
    } catch {
        throw new Error('Could not refresh OneDrive access. Reconnect Personal Slide Publisher and replace ONEDRIVE_AUTH_CACHE; no interactive sign-in is attempted in Actions.');
    }

    // Persist renewal before any download or conversion can fail.
    const renewedCache = client.getTokenCache().serialize();
    if (renewedCache !== savedCache) {
        await saveCache(renewedCache);
        log('Saved renewed OneDrive authentication.');
    }
    return createDriveReader(auth.accessToken, request);
}

export function createDriveReader(accessToken, request = fetch) {
    async function graph(resource, select = '') {
        const url = new URL(`https://graph.microsoft.com/v1.0/me/drive${resource}`);
        if (select) url.searchParams.set('$select', select);
        try {
            return await request(url, {
                headers: { Authorization: `Bearer ${accessToken}` },
                redirect: 'manual',
                signal: AbortSignal.timeout(30_000)
            });
        } catch {
            throw new Error('Microsoft Graph request failed or timed out. Retry the workflow.');
        }
    }

    async function metadata(resource, expectedId) {
        const response = await graph(resource, 'id,name,eTag,size,file,webUrl');
        if (!response.ok) {
            throw new Error(`OneDrive metadata request failed${expectedId ? ` for ${expectedId}` : ''} (HTTP ${response.status}). The deck may be missing or inaccessible.`);
        }
        let item;
        try {
            item = await response.json();
        } catch {
            throw new Error('OneDrive returned an invalid metadata response.');
        }
        if (typeof item?.id !== 'string' || !item.id || (expectedId !== undefined && item.id !== expectedId) ||
            typeof item.name !== 'string' || !/\.pptx$/i.test(item.name) || item.file?.mimeType !== mimeType ||
            typeof item.eTag !== 'string' || !item.eTag || !Number.isSafeInteger(item.size) || item.size <= 0) {
            throw new Error('OneDrive returned invalid PowerPoint metadata for the requested file.');
        }
        return item;
    }

    return {
        getItem(id) {
            return metadata(`/items/${encodeURIComponent(id)}`, id);
        },
        getItemByPath(filePath) {
            return metadata(drivePath(filePath));
        },
        async download(item, destination) {
            let response = await graph(`/items/${encodeURIComponent(item.id)}/content`);
            if (response.status === 302) {
                let location;
                try {
                    location = new URL(response.headers.get('location'));
                    if (location.protocol !== 'https:') throw new Error('Expected HTTPS.');
                } catch {
                    throw new Error('OneDrive returned an invalid download redirect.');
                }
                try {
                    // The signed download URL does not need, and must not receive, the Graph bearer token.
                    response = await request(location, { signal: AbortSignal.timeout(180_000) });
                } catch {
                    throw new Error('PowerPoint download failed or timed out. Retry the workflow.');
                }
            }
            if (!response.ok || !response.body) {
                throw new Error(`PowerPoint download failed (HTTP ${response.status}).`);
            }
            try {
                await pipeline(Readable.fromWeb(response.body), createWriteStream(destination, { flags: 'wx' }));
            } catch {
                throw new Error('Could not finish the temporary PowerPoint download.');
            }
            if ((await stat(destination)).size !== item.size) {
                throw new Error('PowerPoint download size changed or was incomplete. Retry after OneDrive finishes syncing.');
            }
            const current = await this.getItem(item.id);
            if (current.eTag !== item.eTag || current.size !== item.size) {
                throw new Error('The PowerPoint changed during download. Retry after OneDrive finishes syncing.');
            }
        }
    };
}
