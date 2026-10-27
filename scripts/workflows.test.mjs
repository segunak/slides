import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { posix } from 'node:path';
import { test } from 'node:test';
import { runInNewContext } from 'node:vm';
import { parseDocument } from 'yaml';
import { readConfig, readState } from './update-decks.mjs';

const root = new URL('../', import.meta.url);
async function workflow(name) {
    const doc = parseDocument(await readFile(new URL(`.github/workflows/${name}.yml`, root), 'utf8'));
    assert.deepEqual(doc.errors, []);
    return doc.toJS();
}

test('repository deck configuration and last-published list are readable', async () => {
    readConfig(await readFile(new URL('decks.yml', root), 'utf8'));
    readState(await readFile(new URL('.github/decks-state.json', root), 'utf8'));
});

test('updater supports scheduled/manual runs and excludes documentation and generated-only changes', async () => {
    const update = await workflow('update-decks');
    assert.equal(update.on.schedule[0].cron, '17 */6 * * *');
    assert.equal(update.on.workflow_dispatch.inputs.force.type, 'boolean');
    assert.equal(update.on.workflow_dispatch.inputs.force.default, false);
    assert.deepEqual(update.on.push.branches, ['main']);
    assert.equal(update.on.push.paths, undefined);
    assert.equal(update.on.pull_request.paths, undefined);
    assert.deepEqual(update.on.push['paths-ignore'], update.on.pull_request['paths-ignore']);
    const triggers = files => files.some(file =>
        !update.on.push['paths-ignore'].some(pattern =>
            // GitHub's recursive Markdown pattern also matches files inside dot-directories.
            pattern === '**/*.md' ? file.endsWith('.md') : posix.matchesGlob(file, pattern)));
    for (const file of [
        'decks.yml', 'package.json', 'package-lock.json',
        'scripts/update-decks.mjs', 'scripts/sync-decks.test.mjs', 'scripts/nested/helper.mjs',
        '.github/skills/onedrive-deck-id/scripts/deck-id.mjs',
        '.github/skills/onedrive-deck-id/tests/deck-id.test.mjs',
        '.github/workflows/update-decks.yml', '.github/workflows/publish-pages.yml',
        '.github/workflows/unrelated.yml', 'src/new-module.js', 'future-code/helper.js'
    ]) assert.ok(triggers([file]), `Code or configuration must trigger synchronization: ${file}`);
    for (const file of [
        'README.md', 'LICENSE', '.gitignore', 'docs/notes.md',
        '.github/copilot-instructions.md', '.github/instructions/slide-publisher.instructions.md',
        '.github/skills/onedrive-deck-id/SKILL.md',
        'decks/index.html', 'decks/example/assets/image.png', '.github/decks-state.json'
    ]) assert.ok(!triggers([file]), `Documentation or generated-only changes must not trigger synchronization: ${file}`);
    assert.ok(!triggers(['README.md', 'docs/guide.html', 'decks/index.html', '.github/decks-state.json']));
    assert.ok(triggers(['README.md', 'scripts/update-decks.mjs']));
    assert.ok(triggers(['decks/index.html', 'future-code/helper.js']));
    assert.equal(update.concurrency['cancel-in-progress'], false);
});

test('pull requests only run offline tests; live credentials and app permissions stay in the main-branch job', async () => {
    const update = await workflow('update-decks');
    assert.ok(update.on.pull_request);
    assert.match(update.jobs.update.if, /github\.event_name != 'pull_request'/);
    assert.match(update.jobs.update.if, /refs\/heads\/main/);
    assert.equal(update.jobs.test.if, "github.event_name == 'pull_request'");
    assert.equal(update.jobs.update.needs, undefined);
    assert.doesNotMatch(JSON.stringify(update.jobs.test), /secrets\.|vars\.ONEDRIVE/);
    assert.deepEqual(update.permissions, { contents: 'read' });
    const token = update.jobs.update.steps.find(step => step.id === 'publisher');
    assert.equal(token.with['permission-secrets'], 'write');
    assert.equal(token.with['permission-contents'], 'write');
    assert.ok(!('permission-workflows' in token.with));
});

test('updater runs the Actions-only synchronization coordinator with its app token and exact-SHA outputs', async () => {
    const update = await workflow('update-decks');
    const sync = update.jobs.update.steps.find(step => step.id === 'sync');
    assert.equal(sync.run, 'node scripts/sync-decks.mjs');
    assert.equal(sync.env.GH_TOKEN, '${{ steps.publisher.outputs.token }}');
    assert.equal(sync.env.APP_SLUG, '${{ steps.publisher.outputs.app-slug }}');
    assert.equal(update.jobs.update.outputs.changed, '${{ steps.sync.outputs.changed }}');
    assert.equal(update.jobs.update.outputs.sha, '${{ steps.sync.outputs.sha }}');
    const checkout = update.jobs.update.steps.find(step => step.uses?.startsWith('actions/checkout@'));
    assert.equal(checkout.with.ref, 'main');
});

test('pull requests install Chromium and run browser/PDF regressions and upstream verification', async () => {
    const update = await workflow('update-decks');
    assert.equal(update.jobs.test['runs-on'], 'windows-latest');
    assert.equal(update.jobs.update['runs-on'], 'windows-latest');
    assert.deepEqual(update.jobs.test.steps.filter(step => step.run).map(step => step.run), [
        'npm ci --ignore-scripts --no-fund',
        'npx playwright install chromium',
        'npm test',
        'npm run verify:converter'
    ]);
});

test('publication requires successful synchronization and publishes code pushes/manual updates even without generated changes', async () => {
    const update = await workflow('update-decks');
    assert.equal(update.jobs.publish.needs, 'update');
    for (const result of ['success', 'failure', 'cancelled', 'skipped']) {
        for (const event of ['push', 'workflow_dispatch', 'schedule']) {
            for (const changed of ['true', 'false', '']) {
                const actual = runInNewContext(update.jobs.publish.if, {
                    needs: { update: { result, outputs: { changed } } },
                    github: { event_name: event }
                });
                const expected = result === 'success' && (changed === 'true' || event === 'push' || event === 'workflow_dispatch');
                assert.equal(actual, expected, `Unexpected publication decision: ${result}, ${event}, changed=${changed}`);
            }
        }
    }
    assert.equal(update.jobs.publish.uses, './.github/workflows/publish-pages.yml');
    assert.equal(update.jobs.publish.with.commit, '${{ needs.update.outputs.sha }}');
    assert.deepEqual(update.jobs.publish.permissions, { contents: 'read', pages: 'write', 'id-token': 'write' });
    assert.equal(update.jobs.publish.secrets, undefined);
});

test('Pages is reusable and manually dispatchable, but never independently triggered by a push', async () => {
    const pages = await workflow('publish-pages');
    assert.deepEqual(Object.keys(pages.on).sort(), ['workflow_call', 'workflow_dispatch']);
    assert.equal(pages.on.workflow_call.inputs.commit.type, 'string');
    assert.equal(pages.on.workflow_call.inputs.commit.required, true);
    assert.equal(pages.on.workflow_dispatch, null);
    const checkout = pages.jobs.deploy.steps.find(step => step.uses?.startsWith('actions/checkout@'));
    assert.equal(checkout.with.ref, "${{ inputs.commit || 'main' }}");
});

test('manual publishing shares the full-pipeline lock without deadlocking the updater calling Pages', async () => {
    const update = await workflow('update-decks');
    const pages = await workflow('publish-pages');
    assert.match(update.concurrency.group, /slides-publication/);
    assert.match(pages.concurrency.group, /inputs\.commit != ''/);
    assert.match(pages.concurrency.group, /slides-publication-child-\{0\}.*github\.run_id/);
    assert.match(pages.concurrency.group, / \|\| 'slides-publication'/);
    for (const flow of [update, pages]) {
        assert.equal(flow.concurrency['cancel-in-progress'], false);
        assert.equal(flow.concurrency.queue, 'max');
    }
});

test('Pages publishes only decks, rejects superseded output, and needs no Node or Microsoft credentials', async () => {
    const pages = await workflow('publish-pages');
    const deploy = pages.jobs.deploy;
    assert.equal(deploy['runs-on'], 'ubuntu-latest');
    assert.equal(deploy.environment.name, 'github-pages');
    assert.deepEqual(deploy.permissions, { contents: 'read', pages: 'write', 'id-token': 'write' });
    const upload = deploy.steps.find(step => step.uses?.startsWith('actions/upload-pages-artifact@'));
    assert.equal(upload.with.path, 'decks');
    const guard = deploy.steps.find(step => step.name === 'Refuse superseded presentation output');
    assert.match(guard.run, /git fetch --no-tags origin main/);
    assert.match(guard.run, /git rev-parse HEAD:decks/);
    assert.match(guard.run, /git rev-parse origin\/main:decks/);
    assert.match(guard.run, /exit 1/);
    assert.ok(deploy.steps.indexOf(guard) > deploy.steps.indexOf(upload));
    assert.ok(deploy.steps.indexOf(guard) < deploy.steps.findIndex(step => step.id === 'deployment'));
    assert.doesNotMatch(JSON.stringify(pages), /jekyll|msal|ONEDRIVE|pptx2html|npm|setup-node/i);
});

test('all external actions are SHA-pinned and checkouts do not persist credentials', async () => {
    for (const name of ['update-decks', 'publish-pages']) {
        for (const job of Object.values((await workflow(name)).jobs)) {
            for (const step of job.steps || []) {
                if (!step.uses) continue;
                assert.match(step.uses, /^actions\/[a-z-]+@[a-f0-9]{40}$/);
                if (step.uses.startsWith('actions/checkout@')) assert.equal(step.with['persist-credentials'], false);
            }
        }
    }
});
