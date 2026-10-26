import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { PublicClientApplication } from '@azure/msal-node';
import { createDriveReader, drivePath, oneDriveClientConfig } from '../../../../scripts/onedrive.mjs';

const help = `Look up verified PowerPoint file IDs using read-only Microsoft Graph.

  npm run deck-id -- "https://1drv.ms/p/..."
  npm run deck-id -- --path "Documents/Talks/My Deck.pptx"
  npm run deck-id -- "https://1drv.ms/p/..." "https://1drv.ms/p/..."

Paste a sharing link, embed URL, or iframe as a quoted argument.
Use --path for an exact path relative to your OneDrive root.
One Microsoft sign-in covers the entire batch. Tokens stay in memory.
Results are a JSON array of verified id, name, and size values.
For machine-readable stdout, use npm run --silent deck-id -- ...
No files, publishing secrets, or OneDrive contents are modified.
`;

function publicUrl(value, base) {
    let url;
    try {
        url = new URL(value, base);
    } catch {
        throw new Error('Expected a public OneDrive URL. Use --path for a OneDrive-relative file path.');
    }
    if (url.protocol !== 'https:' || url.username || url.password || url.port ||
        !['1drv.ms', 'onedrive.live.com'].includes(url.hostname)) {
        throw new Error('Only HTTPS personal OneDrive links on 1drv.ms or onedrive.live.com are supported. Use --path for other link formats.');
    }
    return url;
}

export async function resolvePublicLink(input, request = fetch) {
    const source = input.trim();
    const iframe = source.match(/(?:data-speaking-talk-slides-src|src)\s*=\s*["']([^"']+)["']/i);
    const value = (iframe ? iframe[1] : source).replace(/&(?:amp|#0*38|#x0*26);/gi, '&');
    let url = publicUrl(value);
    for (let redirects = 0; redirects <= 8; redirects++) {
        if (url.hostname === 'onedrive.live.com') {
            const ids = [...url.searchParams.getAll('id'), ...url.searchParams.getAll('resid')].filter(Boolean);
            if (ids.length) {
                if (new Set(ids).size !== 1) throw new Error('The OneDrive link contains conflicting file IDs. Use --path to look up the file directly.');
                return ids[0];
            }
        }
        if (redirects === 8) break;
        let response;
        try {
            // Public-link resolution must never receive the Microsoft Graph bearer token.
            response = await request(url, { redirect: 'manual', signal: AbortSignal.timeout(30_000) });
            await response.body?.cancel();
        } catch {
            throw new Error('Could not expand the public OneDrive link. Retry or use --path with its exact OneDrive path.');
        }
        if (![301, 302, 303, 307, 308].includes(response.status) || !response.headers.get('location')) {
            throw new Error('The public link did not expose a file ID. Use --path with its exact OneDrive path; no ID was guessed.');
        }
        url = publicUrl(response.headers.get('location'), url);
    }
    throw new Error('The public link exceeded eight redirects. Use --path with its exact OneDrive path.');
}

function readClientId() {
    try {
        return execFileSync('gh', ['variable', 'get', 'ONEDRIVE_CLIENT_ID', '--repo', 'segunak/slides'], {
            encoding: 'utf8', stdio: ['pipe', 'pipe', 'pipe'], timeout: 30_000, windowsHide: true
        }).trim();
    } catch {
        throw new Error('Set ONEDRIVE_CLIENT_ID in your environment, or sign into GitHub CLI so it can read the slides repository variable.');
    }
}

export async function connectLookup({
    env = process.env,
    clientIdFromGitHub = readClientId,
    createClient = config => new PublicClientApplication(config),
    request = fetch,
    report = console.error
} = {}) {
    const clientId = env.ONEDRIVE_CLIENT_ID?.trim() || clientIdFromGitHub();
    if (!/^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/i.test(clientId)) {
        throw new Error('ONEDRIVE_CLIENT_ID must be the Microsoft Application (client) ID.');
    }
    const client = createClient(oneDriveClientConfig(clientId));
    let auth;
    if (env.ONEDRIVE_AUTH_CACHE) {
        try {
            client.getTokenCache().deserialize(env.ONEDRIVE_AUTH_CACHE);
            const accounts = await client.getTokenCache().getAllAccounts();
            if (accounts.length !== 1) throw new Error('Expected one personal account.');
            auth = await client.acquireTokenSilent({ account: accounts[0], scopes: ['Files.Read'] });
        } catch {
            throw new Error('The provided OneDrive cache could not sign in silently. For local lookup, remove ONEDRIVE_AUTH_CACHE from the environment and rerun to sign in.');
        }
    } else {
        if (env.GITHUB_ACTIONS === 'true' || (env.CI && env.CI !== 'false')) {
            throw new Error('Interactive lookup is disabled in CI. Supply ONEDRIVE_AUTH_CACHE securely, or run this command locally.');
        }
        report('Sign in with the personal Microsoft account that owns the decks. Access is read-only and credentials are not saved.');
        try {
            auth = await client.acquireTokenByDeviceCode({
                scopes: ['Files.Read'],
                deviceCodeCallback: response => report(response.message)
            });
        } catch {
            throw new Error('Microsoft sign-in failed or expired. Confirm public client flows are enabled and retry the command.');
        }
    }
    if (!auth?.accessToken) throw new Error('Microsoft did not return a read-only access token. No IDs were verified.');
    return createDriveReader(auth.accessToken, request);
}

export async function lookupDecks(references, drive) {
    const results = new Map();
    for (const reference of references) {
        const first = reference.path !== undefined
            ? await drive.getItemByPath(reference.path)
            : await drive.getItem(reference.id);
        const verified = await drive.getItem(first.id);
        if (first.id !== verified.id || first.name !== verified.name || first.eTag !== verified.eTag || first.size !== verified.size) {
            throw new Error('The file changed during verification. Let OneDrive finish syncing and retry.');
        }
        results.set(verified.id, { id: verified.id, name: verified.name, size: verified.size });
    }
    return [...results.values()];
}

export async function runCli(args, {
    resolve = resolvePublicLink,
    connect = connectLookup,
    write = value => process.stdout.write(value)
} = {}) {
    const { values, positionals } = parseArgs({
        args, allowPositionals: true,
        options: { path: { type: 'string', multiple: true }, help: { type: 'boolean', short: 'h' } }
    });
    if (values.help) {
        write(help);
        return;
    }
    if (!positionals.length && !values.path?.length) {
        throw new Error('Provide a public OneDrive link or --path. Run npm run deck-id -- --help for examples.');
    }
    const references = [];
    for (const filePath of values.path || []) {
        drivePath(filePath);
        references.push({ path: filePath });
    }
    for (const link of positionals) references.push({ id: await resolve(link) });
    const results = await lookupDecks(references, await connect());
    write(JSON.stringify(results, null, 2) + '\n');
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    runCli(process.argv.slice(2)).catch(error => {
        console.error(`Deck ID lookup failed: ${error.message}`);
        process.exitCode = 1;
    });
}
