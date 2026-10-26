import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { test } from 'node:test';
import { syncDecks, writeSyncOutputs } from './sync-decks.mjs';

const first = 'a'.repeat(40);
const newer = 'b'.repeat(40);
const generated = 'c'.repeat(40);
const regenerated = 'd'.repeat(40);

function fixture({ changes = [true], remote = first, checkout = first, push, afterPush, afterUpdate, failure } = {}) {
    const calls = [];
    const logs = [];
    let head = checkout;
    let fetched = remote;
    let attempt = 0;
    let commits = 0;
    const options = {
        env: { GITHUB_ACTIONS: 'true', GH_TOKEN: 'app-token', APP_SLUG: 'slide-publisher' },
        log: message => logs.push(message),
        run(command, args, settings = {}) {
            calls.push({ command, args, settings, head });
            const operation = [command, ...args].join(' ');
            if (failure === operation) throw new Error(`Failed: ${operation}`);
            const result = (stdout = '', status = 0) => ({ stdout, status });
            if (command === 'gh') return result(args[0] === 'api' ? '1234\n' : '');
            if (command === 'npm') {
                if (args[0] === 'run') afterUpdate?.({ attempt, advance: sha => { remote = sha; } });
                return result();
            }
            assert.equal(command, 'git');
            switch (args[0]) {
                case 'status': return result();
                case 'config': return result();
                case 'fetch':
                    fetched = remote;
                    return result();
                case 'switch':
                    assert.deepEqual(args, ['switch', '--detach', 'origin/main']);
                    head = fetched;
                    attempt++;
                    return result();
                case 'rev-parse':
                    return result(args[1] === 'HEAD' ? head : fetched);
                case 'add':
                    assert.deepEqual(args, ['add', '--', 'decks', '.github/decks-state.json']);
                    return result();
                case 'diff':
                    return result('', (changes[attempt - 1] ?? true) ? 1 : 0);
                case 'commit':
                    head = ++commits === 1 ? generated : regenerated;
                    return result();
                case 'push': {
                    const rejected = push?.({ attempt, head, advance: sha => { remote = sha; } });
                    if (rejected) return result('', 1);
                    remote = head;
                    afterPush?.({ advance: sha => { remote = sha; } });
                    return result();
                }
                default: assert.fail(`Unexpected Git operation: ${operation}`);
            }
        }
    };
    return { options, calls, logs };
}

test('commits only generated output after checking out, installing, testing, and converting latest main', () => {
    const f = fixture({ remote: newer });
    assert.deepEqual(syncDecks(f.options), { changed: true, sha: generated });
    const commands = f.calls.map(call => [call.command, ...call.args].join(' '));
    const ordered = [
        'git fetch --no-tags origin main',
        'git switch --detach origin/main',
        'npm ci --ignore-scripts --no-fund',
        'npm test',
        'npm run update-decks',
        'git add -- decks .github/decks-state.json',
        'git commit -m Update published decks',
        'git push origin HEAD:main'
    ];
    let position = -1;
    for (const command of ordered) {
        position = commands.indexOf(command, position + 1);
        assert.ok(position >= 0, `Missing or out-of-order operation: ${command}`);
    }
    assert.equal(f.calls.find(call => call.command === 'npm').head, newer);
    assert.deepEqual(f.calls.filter(call => call.command === 'git' && call.args[0] === 'push').map(call => call.args),
        [['push', 'origin', 'HEAD:main']]);
    assert.doesNotMatch(commands.join('\n'), /--force|reset|restore|rebase|git pull/);
});

test('unchanged output returns no-publication without committing or pushing', () => {
    const f = fixture({ changes: [false] });
    assert.deepEqual(syncDecks(f.options), { changed: false, sha: first });
    assert.ok(!f.calls.some(call => ['commit', 'push'].includes(call.args[0])));
});

test('a concurrent user push causes a clean retry of latest source rather than merging stale generated files', () => {
    const f = fixture({
        changes: [true, true],
        push: ({ attempt, advance }) => {
            if (attempt !== 1) return false;
            advance(newer);
            return true;
        }
    });
    assert.deepEqual(syncDecks(f.options), { changed: true, sha: regenerated });
    const conversions = f.calls.filter(call => call.command === 'npm' && call.args[0] === 'run');
    assert.deepEqual(conversions.map(call => call.head), [first, newer]);
    assert.equal(f.calls.filter(call => call.command === 'npm' && call.args[0] === 'test').length, 2);
    assert.equal(f.logs.length, 1);
    assert.match(f.logs[0], /main changed.*Retrying/);
});

test('an initially unchanged run retries if main changes while it checks the decks', () => {
    const f = fixture({
        changes: [false, true],
        afterUpdate: ({ attempt, advance }) => { if (attempt === 1) advance(newer); }
    });
    assert.deepEqual(syncDecks(f.options), { changed: true, sha: generated });
    assert.equal(f.calls.filter(call => call.command === 'npm' && call.args[0] === 'run').length, 2);
});

test('a push failure with unchanged main stops instead of retrying permissions or publication failures', () => {
    const f = fixture({ push: () => true });
    assert.throws(() => syncDecks(f.options), /could not be pushed.*nothing will be published/);
    assert.equal(f.calls.filter(call => call.args[0] === 'push').length, 1);
});

test('a reported push error is accepted only when the generated commit is actually confirmed on main', () => {
    const f = fixture({ push: ({ head, advance }) => { advance(head); return true; } });
    assert.deepEqual(syncDecks(f.options), { changed: true, sha: generated });
    assert.match(f.logs[0], /confirmed on main/);
});

test('does not lose a successful publication handoff when another source commit arrives after the push', () => {
    const f = fixture({
        changes: [true, false],
        afterPush: ({ advance }) => advance(newer)
    });
    assert.deepEqual(syncDecks(f.options), { changed: true, sha: generated });
    assert.equal(f.calls.filter(call => call.command === 'npm' && call.args[0] === 'run').length, 1);
});

test('repeated branch movement is bounded and never produces publication outputs', () => {
    let moves = 0;
    const f = fixture({
        push: ({ advance }) => { advance(String(++moves).repeat(40)); return true; }
    });
    assert.throws(() => syncDecks(f.options), /main kept changing.*No outdated output/);
    assert.equal(f.calls.filter(call => call.args[0] === 'push').length, 3);
    assert.equal(f.logs.length, 2);
});

for (const failure of ['npm ci --ignore-scripts --no-fund', 'npm test', 'npm run update-decks']) {
    test(`${failure} failure cannot commit or publish output`, () => {
        const f = fixture({ failure });
        assert.throws(() => syncDecks(f.options), /Failed:/);
        assert.ok(!f.calls.some(call => call.command === 'git' && ['add', 'commit', 'push'].includes(call.args[0])));
    });
}

test('requires Actions and app credentials before running any Git command', () => {
    for (const env of [{}, { GITHUB_ACTIONS: 'true' }, { GITHUB_ACTIONS: 'true', GH_TOKEN: 'token', APP_SLUG: 'bad/slug' }]) {
        assert.throws(() => syncDecks({ env, run: () => assert.fail('No commands expected') }),
            /only allowed in GitHub Actions|Missing GitHub App/);
    }
});

test('does not overwrite a dirty checkout', () => {
    const f = fixture();
    f.options.run = () => ({ status: 0, stdout: ' M README.md\n' });
    assert.throws(() => syncDecks(f.options), /must be clean/);
});

test('writes exact-SHA workflow outputs only from a valid successful synchronization result', async t => {
    const directory = await mkdtemp(path.join(tmpdir(), 'slides-sync-outputs-'));
    t.after(() => rm(directory, { recursive: true, force: true }));
    const file = path.join(directory, 'output');
    await writeSyncOutputs(file, { changed: true, sha: generated });
    await writeSyncOutputs(file, { changed: false, sha: first });
    assert.equal(await readFile(file, 'utf8'),
        `changed=true\nsha=${generated}\nchanged=false\nsha=${first}\n`);
    const before = await readFile(file, 'utf8');
    for (const result of [{ changed: 'true', sha: first }, { changed: true, sha: 'main' }]) {
        await assert.rejects(writeSyncOutputs(file, result), /invalid workflow/);
    }
    assert.equal(await readFile(file, 'utf8'), before);
});
