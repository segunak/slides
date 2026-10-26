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

test('updater schedules four checks and excludes its generated commits from update triggers', async () => {
    const update = await workflow('update-decks');
    assert.equal(update.on.schedule[0].cron, '17 */6 * * *');
    assert.equal(update.on.workflow_dispatch.inputs.force.type, 'boolean');
    assert.equal(update.on.workflow_dispatch.inputs.force.default, false);
    assert.deepEqual(update.on.push.branches, ['main']);
    assert.ok(!update.on.push.paths.some(file => file.startsWith('decks/') || file === '.github/decks-state.json'));
    for (const input of ['decks.yml', 'package-lock.json', 'scripts/**']) assert.ok(update.on.push.paths.includes(input));
    assert.equal(update.concurrency['cancel-in-progress'], false);
});

test('pull requests only run offline tests; live credentials and app permissions stay in the main-branch job', async () => {
    const update = await workflow('update-decks');
    assert.ok(update.on.pull_request);
    assert.ok(update.on.pull_request.paths.includes('.github/skills/onedrive-deck-id/**'));
    assert.match(update.jobs.update.if, /github\.event_name != 'pull_request'/);
    assert.match(update.jobs.update.if, /refs\/heads\/main/);
    assert.equal(update.jobs.update.needs, 'test');
    assert.doesNotMatch(JSON.stringify(update.jobs.test), /secrets\.|vars\.ONEDRIVE/);
    assert.deepEqual(update.permissions, { contents: 'read' });
    const token = update.jobs.update.steps.find(step => step.id === 'publisher');
    assert.equal(token.with['permission-secrets'], 'write');
    assert.equal(token.with['permission-contents'], 'write');
    assert.ok(!('permission-workflows' in token.with));
});

test('automatic commits stage only generated paths and use the GitHub App token without force pushes', async () => {
    const update = await workflow('update-decks');
    const commit = update.jobs.update.steps.find(step => step.name === 'Commit generated changes');
    assert.match(commit.run, /git add -- decks \.github\/decks-state\.json/);
    assert.doesNotMatch(commit.run, /git add (?:\.|-A)|--force|git pull|git rebase/);
    assert.match(commit.run, /git diff --cached --quiet/);
    assert.match(commit.run, /git push origin HEAD:main/);
    assert.equal(commit.env.GH_TOKEN, '${{ steps.publisher.outputs.token }}');
});

test('Pages publishes only decks and needs no Node, Microsoft credentials, or converter', async () => {
    const pages = await workflow('publish-pages');
    assert.ok(pages.on.push.paths.includes('decks/**'));
    assert.deepEqual(pages.on.push.branches, ['main']);
    const deploy = pages.jobs.deploy;
    assert.equal(deploy.environment.name, 'github-pages');
    assert.deepEqual(deploy.permissions, { contents: 'read', pages: 'write', 'id-token': 'write' });
    const upload = deploy.steps.find(step => step.uses?.startsWith('actions/upload-pages-artifact@'));
    assert.equal(upload.with.path, 'decks');
    assert.doesNotMatch(JSON.stringify(pages), /jekyll|msal|ONEDRIVE|pptx2html|npm|setup-node/i);
});

test('all external actions are SHA-pinned and checkouts do not persist credentials', async () => {
    for (const name of ['update-decks', 'publish-pages']) {
        for (const job of Object.values((await workflow(name)).jobs)) {
            for (const step of job.steps) {
                if (!step.uses) continue;
                assert.match(step.uses, /^actions\/[a-z-]+@[a-f0-9]{40}$/);
                if (step.uses.startsWith('actions/checkout@')) assert.equal(step.with['persist-credentials'], false);
            }
        }
    }
});
