'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { test } = require('node:test');
const { findSignedArtifact } = require('./gallery.js');

const SHA = 'a'.repeat(40);
const OTHER_SHA = 'b'.repeat(40);
const REPOSITORY_ID = 12345;

function artifact(id = 101, changes = {}) {
  return {
    id, name: 'AxialSqlTools-Signed-4.14', expired: false, size_in_bytes: 1234,
    workflow_run: {
      id: id + 1000, repository_id: REPOSITORY_ID, head_repository_id: REPOSITORY_ID,
      head_sha: SHA, head_branch: 'main'
    },
    ...changes
  };
}

function run(id = 1101, changes = {}) {
  return {
    id, path: '.github/workflows/build-sign-vsix.yml', head_sha: SHA,
    event: 'workflow_dispatch', status: 'completed', conclusion: 'success',
    repository: { id: REPOSITORY_ID, full_name: 'Axial-SQL/AxialSqlTools' },
    head_repository: { id: REPOSITORY_ID, full_name: 'Axial-SQL/AxialSqlTools' },
    html_url: `https://github.com/Axial-SQL/AxialSqlTools/actions/runs/${id}`,
    ...changes
  };
}

function fixture(t, { artifacts = [], runs = new Map(), context: contextChanges = {} } = {}) {
  const repoRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'axial-gallery-'));
  t.after(() => fs.rmSync(repoRoot, { recursive: true, force: true }));
  fs.mkdirSync(path.join(repoRoot, 'AxialSqlTools/Properties'), { recursive: true });
  fs.writeFileSync(path.join(repoRoot, 'AxialSqlTools/source.extension.vsixmanifest'), '<Identity Version="4.14" />');
  fs.writeFileSync(path.join(repoRoot, 'AxialSqlTools/Properties/AssemblyInfo.cs'),
    '[assembly: AssemblyVersion("4.14.0.0")]\n[assembly: AssemblyFileVersion("4.14.0.0")]\n');
  const context = {
    repo: { owner: 'Axial-SQL', repo: 'AxialSqlTools' },
    payload: { repository: { id: REPOSITORY_ID, default_branch: 'main' } },
    eventName: 'workflow_dispatch', ref: 'refs/heads/main', sha: SHA, runId: 9000,
    ...contextChanges
  };
  const outputs = {};
  const summaries = [];
  const listCalls = [];
  const runCalls = [];
  const core = {
    setOutput: (key, value) => { outputs[key] = value; },
    summary: { addRaw: value => ({ write: async () => { summaries.push(value); } }) }
  };
  const github = { rest: { actions: {
    listArtifactsForRepo: async parameters => {
      listCalls.push(parameters);
      const start = (parameters.page - 1) * parameters.per_page;
      return { data: { artifacts: artifacts.slice(start, start + parameters.per_page), total_count: artifacts.length } };
    },
    getWorkflowRun: async parameters => {
      runCalls.push(parameters);
      const result = runs.get(parameters.run_id);
      if (result instanceof Error) throw result;
      if (!result) throw Object.assign(new Error('Run not found'), { status: 404 });
      return { data: result };
    }
  } } };
  return { github, context, core, repoRoot, outputs, summaries, listCalls, runCalls };
}

test('selects the newest signed artifact by exact name and commit', async t => {
  const f = fixture(t, {
    artifacts: [artifact(101), artifact(200, { name: 'unsigned-vsix' }), artifact(103), artifact(102),
      artifact(300, { workflow_run: { id: 1300, head_sha: OTHER_SHA } })],
    runs: new Map([[1101, run(1101)], [1102, run(1102)], [1103, run(1103)]])
  });
  await findSignedArtifact(f);
  assert.deepEqual(f.outputs, { version: '4.14', source_sha: SHA, signed_artifact_id: '103', signed_run_id: '1103' });
  assert.deepEqual(f.listCalls, [{ owner: 'Axial-SQL', repo: 'AxialSqlTools', name: 'AxialSqlTools-Signed-4.14', per_page: 100, page: 1 }]);
  assert.deepEqual(f.runCalls.map(call => call.run_id), [1103]);
  assert.match(f.summaries[0], /Reuse artifact 103/);
});

test('ignores expired, empty, forked, and missing-run artifacts', async t => {
  const f = fixture(t, {
    artifacts: [artifact(108, { expired: true }), artifact(107, { size_in_bytes: 0 }),
      artifact(106, { workflow_run: { id: 1106, head_sha: SHA, repository_id: 999 } }),
      artifact(105, { workflow_run: { id: 1105, head_sha: SHA, head_repository_id: 999 } }),
      artifact(104, { workflow_run: null }), artifact(101)],
    runs: new Map([[1101, run()]])
  });
  await findSignedArtifact(f);
  assert.equal(f.outputs.signed_artifact_id, '101');
  assert.deepEqual(f.runCalls.map(call => call.run_id), [1101]);
});

test('checks the full run path, event, commit, and repository before reuse', async t => {
  const changes = [
    { path: '.github/workflows/untrusted.yml' },
    { path: 'build-sign-vsix.yml' },
    { event: 'pull_request' },
    { head_sha: OTHER_SHA },
    { repository: { id: 999, full_name: 'Axial-SQL/AxialSqlTools' } },
    { head_repository: { id: 999, full_name: 'someone/AxialSqlTools' } },
    { repository: { full_name: 'someone/AxialSqlTools' } },
    { head_repository: { full_name: 'someone/AxialSqlTools' } }
  ];
  const artifacts = changes.map((_, index) => artifact(110 - index));
  const runs = new Map(changes.map((value, index) => [1110 - index, run(1110 - index, value)]));
  artifacts.push(artifact(101));
  runs.set(1101, run());
  const f = fixture(t, { artifacts, runs });
  await findSignedArtifact(f);
  assert.equal(f.outputs.signed_artifact_id, '101');
  assert.equal(f.runCalls.length, changes.length + 1);
});

test('reuses a verified artifact when downstream release or gallery publishing failed', async t => {
  for (const workflow of ['prepare-new-release.yml', 'publish-to-galleries.yml']) {
    await t.test(workflow, async st => {
      const f = fixture(st, {
        artifacts: [artifact()],
        runs: new Map([[1101, run(1101, { path: `.github/workflows/${workflow}`, conclusion: 'failure' })]])
      });
      await findSignedArtifact(f);
      assert.equal(f.outputs.signed_artifact_id, '101');
    });
  }
});

test('reuses an artifact from the current run while retrying publishing', async t => {
  const f = fixture(t, {
    artifacts: [artifact(101, { workflow_run: { id: 9000, head_sha: SHA } })],
    runs: new Map([[9000, run(9000, { path: '.github/workflows/publish-to-galleries.yml', status: 'in_progress', conclusion: null })]])
  });
  await findSignedArtifact(f);
  assert.equal(f.outputs.signed_run_id, '9000');
});

test('permits a manual tag build for the same commit and repository', async t => {
  const f = fixture(t, {
    artifacts: [artifact(101, { workflow_run: { id: 1101, head_sha: SHA, head_branch: '4.14.20261003.2010' } })],
    runs: new Map([[1101, run(1101, {
      head_branch: '4.14.20261003.2010',
      repository: { full_name: 'axial-sql/axialsqltools' },
      head_repository: { id: String(REPOSITORY_ID) }
    })]])
  });
  await findSignedArtifact(f);
  assert.equal(f.outputs.signed_artifact_id, '101');
});

test('accepts the optional workflow ref suffix returned by GitHub', async t => {
  for (const suffix of ['@main', '@refs/tags/4.14.20261003.2010']) {
    const f = fixture(t, {
      artifacts: [artifact()],
      runs: new Map([[1101, run(1101, { path: `.github/workflows/build-sign-vsix.yml${suffix}` })]])
    });
    await findSignedArtifact(f);
    assert.equal(f.outputs.signed_artifact_id, '101');
  }
});

test('reads all artifact pages and selects the newest eligible ID across them', async t => {
  const artifacts = Array.from({ length: 100 }, (_, index) => artifact(200 + index, { expired: true }));
  artifacts[0] = artifact(101);
  artifacts.push(artifact(500), artifact(102));
  const f = fixture(t, { artifacts, runs: new Map([[1101, run(1101)], [1500, run(1500)]]) });
  await findSignedArtifact(f);
  assert.deepEqual(f.listCalls.map(call => call.page), [1, 2]);
  assert.equal(f.outputs.signed_artifact_id, '500');
  assert.deepEqual(f.runCalls.map(call => call.run_id), [1500]);
});

test('rebuilds when there is no reusable artifact', async t => {
  const f = fixture(t, { artifacts: [artifact(101, { expired: true })] });
  await findSignedArtifact(f);
  assert.deepEqual(f.outputs, { version: '4.14', source_sha: SHA, signed_artifact_id: '', signed_run_id: '' });
  assert.equal(f.runCalls.length, 0);
  assert.match(f.summaries[0], /Build and Sign VSIX will prepare one/);
});

test('skips deleted runs and caches rejected run lookups', async t => {
  const f = fixture(t, {
    artifacts: [artifact(104), artifact(103, { workflow_run: { id: 1104, head_sha: SHA } }), artifact(101)],
    runs: new Map([[1101, run()]])
  });
  await findSignedArtifact(f);
  assert.equal(f.outputs.signed_artifact_id, '101');
  assert.deepEqual(f.runCalls.map(call => call.run_id), [1104, 1101]);
});

test('propagates authorization failures instead of silently requesting another build', async t => {
  const error = Object.assign(new Error('Resource not accessible by integration'), { status: 403 });
  const f = fixture(t, { artifacts: [artifact()], runs: new Map([[1101, error]]) });
  await assert.rejects(findSignedArtifact(f), error);
  assert.deepEqual(f.outputs, {});
});

test('requires manual dispatch from the default branch before querying artifacts', async t => {
  for (const changes of [{ eventName: 'push' }, { ref: 'refs/heads/feature' }, { ref: 'refs/tags/4.14' }]) {
    const f = fixture(t, { context: changes });
    await assert.rejects(findSignedArtifact(f), /must be run manually from main/);
    assert.equal(f.listCalls.length, 0);
  }
});

test('uses the repository default branch instead of assuming main', async t => {
  const f = fixture(t, { context: {
    payload: { repository: { id: REPOSITORY_ID, default_branch: 'master' } }, ref: 'refs/heads/master'
  } });
  await findSignedArtifact(f);
  assert.equal(f.outputs.signed_artifact_id, '');
});

test('fails on an inconsistent source version before querying artifacts', async t => {
  const f = fixture(t);
  fs.writeFileSync(path.join(f.repoRoot, 'AxialSqlTools/Properties/AssemblyInfo.cs'),
    '[assembly: AssemblyVersion("4.13.0.0")]\n[assembly: AssemblyFileVersion("4.14.0.0")]\n');
  await assert.rejects(findSignedArtifact(f), /AssemblyVersion must be 4.14.0.0/);
  assert.equal(f.listCalls.length, 0);
});
