import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { test } from 'node:test';
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

test('updater supports scheduled/manual runs and ignores only its generated commits', async () => {
    const update = await workflow('update-decks');
    assert.equal(update.on.schedule[0].cron, '17 */6 * * *');
    assert.equal(update.on.workflow_dispatch.inputs.force.type, 'boolean');
    assert.equal(update.on.workflow_dispatch.inputs.force.default, false);
    assert.deepEqual(update.on.push.branches, ['main']);
    assert.deepEqual(update.on.push['paths-ignore'], ['decks/**', '.github/decks-state.json']);
    assert.equal(update.on.push.paths, undefined);
    assert.equal(update.concurrency['cancel-in-progress'], false);
});

test('pull requests only run offline tests; live credentials and app permissions stay in the main-branch job', async () => {
    const update = await workflow('update-decks');
    assert.ok(update.on.pull_request);
    assert.ok(update.on.pull_request.paths.includes('.github/skills/onedrive-deck-id/**'));
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

test('publication depends on a changed, successfully pushed update and passes its exact SHA', async () => {
    const update = await workflow('update-decks');
    assert.equal(update.jobs.publish.needs, 'update');
    assert.equal(update.jobs.publish.if, "needs.update.outputs.changed == 'true'");
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
