import { spawnSync } from 'node:child_process';
import { appendFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

function runCommand(command, args, { allowFailure = false, inherit = false } = {}) {
    const result = spawnSync(command, args, {
        encoding: 'utf8',
        stdio: inherit ? 'inherit' : ['ignore', 'pipe', 'pipe'],
        windowsHide: true
    });
    if (result.error) throw new Error(`Could not run ${command}: ${result.error.message}`);
    if (result.signal || result.status === null) throw new Error(`${command} was interrupted.`);
    if (result.status !== 0 && !allowFailure) {
        throw new Error(`${command} ${args[0]} failed: ${result.stderr?.trim() || `exit ${result.status}`}`);
    }
    return { status: result.status, stdout: result.stdout || '' };
}

export function syncDecks({ env = process.env, run = runCommand, log = console.log } = {}) {
    if (env.GITHUB_ACTIONS !== 'true') {
        throw new Error('Generated commits are only allowed in GitHub Actions. Use npm run update-decks for local preparation without committing.');
    }
    if (!env.GH_TOKEN || !/^[a-z0-9-]+$/i.test(env.APP_SLUG || '')) {
        throw new Error('Missing GitHub App token or app slug.');
    }
    const git = (...args) => run('git', args).stdout.trim();
    if (git('status', '--porcelain')) throw new Error('The Actions checkout must be clean before synchronization.');
    const botId = run('gh', ['api', `/users/${env.APP_SLUG}[bot]`, '--jq', '.id']).stdout.trim();
    if (!/^\d+$/.test(botId)) throw new Error('GitHub did not return the publisher bot ID.');
    git('config', 'user.name', `${env.APP_SLUG}[bot]`);
    git('config', 'user.email', `${botId}+${env.APP_SLUG}[bot]@users.noreply.github.com`);
    run('gh', ['auth', 'setup-git']);

    for (let attempt = 1; attempt <= 3; attempt++) {
        git('fetch', '--no-tags', 'origin', 'main');
        // A clean detached checkout discards no work and avoids replaying generated files over newer source edits.
        git('switch', '--detach', 'origin/main');
        const base = git('rev-parse', 'HEAD');
        run('npm', ['ci', '--ignore-scripts', '--no-fund'], { inherit: true });
        run('npm', ['test'], { inherit: true });
        run('npm', ['run', 'update-decks'], { inherit: true });
        git('add', '--', 'decks', '.github/decks-state.json');
        const diff = run('git', ['diff', '--cached', '--quiet'], { allowFailure: true });
        if (![0, 1].includes(diff.status)) throw new Error('Could not check the generated changes.');

        let sha = base;
        let push;
        if (diff.status === 1) {
            git('commit', '-m', 'Update published decks');
            sha = git('rev-parse', 'HEAD');
            push = run('git', ['push', 'origin', 'HEAD:main'], { allowFailure: true, inherit: true });
            if (push.status === 0) return { changed: true, sha };
        }

        git('fetch', '--no-tags', 'origin', 'main');
        const latest = git('rev-parse', 'origin/main');
        if (latest === sha) {
            if (push?.status) log('The push reported an error, but the generated commit was confirmed on main.');
            return { changed: diff.status === 1, sha };
        }
        if (latest === base) {
            throw new Error('The generated commit could not be pushed. Check the GitHub App permissions and repository rules; nothing will be published.');
        }
        if (attempt < 3) {
            log(`main changed during synchronization. Retrying from its latest revision (attempt ${attempt + 1}/3).`);
        }
    }
    throw new Error('main kept changing during synchronization. No outdated output will be published; rerun Update Decks after the pushes finish.');
}

export async function writeSyncOutputs(file, result) {
    if (!file || typeof result.changed !== 'boolean' || !/^[a-f0-9]{40}$/.test(result.sha)) {
        throw new Error('Missing or invalid workflow publication outputs.');
    }
    await appendFile(file, `changed=${result.changed}\nsha=${result.sha}\n`);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    try {
        if (!process.env.GITHUB_OUTPUT) throw new Error('Missing GitHub Actions output file.');
        await writeSyncOutputs(process.env.GITHUB_OUTPUT, syncDecks());
    } catch (error) {
        console.error(`Deck synchronization failed: ${error.message}`);
        process.exitCode = 1;
    }
}
