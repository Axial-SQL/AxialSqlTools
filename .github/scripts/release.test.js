'use strict';

const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const { prepare, publish, productVersion, tagForRun } = require('./release');

const sha = 'a'.repeat(40);
const version = '4.15';
const createdAt = '2026-10-03T20:12:59Z';
const tag = '4.15.20261003.2012';
const runId = 12345;
const vsixName = `AxialSqlTools_SSMS22_${version}.vsix`;
const checksumName = `${vsixName}.sha256`;
const digest = data => crypto.createHash('sha256').update(data).digest('hex');
const notFound = () => { throw Object.assign(new Error('Not found'), { status: 404 }); };
const copy = value => structuredClone(value);

function fixture(t, options = {}) {
  const repoRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'axial-release-test-'));
  t.after(() => fs.rmSync(repoRoot, { recursive: true, force: true }));
  fs.mkdirSync(path.join(repoRoot, 'AxialSqlTools/Properties'), { recursive: true });
  fs.mkdirSync(path.join(repoRoot, 'release-assets'));
  fs.writeFileSync(path.join(repoRoot, 'AxialSqlTools/source.extension.vsixmanifest'),
    `<PackageManifest><Metadata><Identity Id="axial" Version="${version}" /></Metadata></PackageManifest>`);
  fs.writeFileSync(path.join(repoRoot, 'AxialSqlTools/Properties/AssemblyInfo.cs'),
    `[assembly: AssemblyVersion("${version}.0.0")]\n[assembly: AssemblyFileVersion("${version}.0.0")]\n`);
  const bytes = Buffer.from('fixture representing the already verified signed VSIX');
  fs.writeFileSync(path.join(repoRoot, 'release-assets', vsixName), bytes);
  fs.writeFileSync(path.join(repoRoot, 'release-assets', checksumName), `${digest(bytes).toUpperCase()}  ${vsixName}\r\n`);

  const context = {
    repo: { owner: 'Axial-SQL', repo: 'AxialSqlTools' },
    sha, runId, ref: 'refs/heads/main', eventName: 'workflow_dispatch',
    serverUrl: 'https://github.com', payload: { repository: { default_branch: 'main' } }
  };
  const outputs = {};
  const summaries = [];
  const core = {
    setOutput: (name, value) => { outputs[name] = value; },
    summary: { addRaw(text) { summaries.push(text); return this; }, async write() {} }
  };
  const env = {
    RELEASE_TAG: tag, RELEASE_SOURCE_SHA: sha, RELEASE_VERSION: version,
    SIGNED_SOURCE_SHA: sha, SIGNED_VERSION: version, SIGNED_VSIX_SHA256: digest(bytes)
  };
  const annotation = {
    sha: 'b'.repeat(40), tag, object: { type: 'commit', sha },
    message: `Prepare New Release\nRun: ${runId}\nSource: ${sha}\nVersion: ${version}`
  };
  const state = {
    calls: [], nextAssetId: 1, latest: { tag_name: '4.14', draft: false },
    run: { head_sha: sha, created_at: createdAt }, annotation: null, ref: null, release: null,
    ...options
  };
  const record = (method, implementation) => async args => {
    state.calls.push({ method, args: copy(args) });
    return { data: copy(await implementation(args)) };
  };
  const rest = {
    actions: { getWorkflowRun: record('getWorkflowRun', () => state.run) },
    git: {
      getRef: record('getRef', () => state.ref || notFound()),
      getTag: record('getTag', () => state.annotation || notFound()),
      createTag: record('createTag', args => {
        state.annotation = { sha: annotation.sha, tag: args.tag, message: args.message,
          object: { sha: args.object, type: args.type } };
        return state.annotation;
      }),
      createRef: record('createRef', args => {
        if (state.ref) throw Object.assign(new Error('Reference already exists'), { status: 422 });
        state.ref = { ref: args.ref, object: { type: 'tag', sha: args.sha } };
        return state.ref;
      })
    },
    repos: {
      getLatestRelease: record('getLatestRelease', () => state.latest || notFound()),
      getReleaseByTag: record('getReleaseByTag', args =>
        state.release?.tag_name === args.tag ? state.release : notFound()),
      generateReleaseNotes: record('generateReleaseNotes', () => ({ body: 'Generated release notes.' })),
      createRelease: record('createRelease', args => {
        state.release = { ...args, id: 99, assets: [], html_url: `https://github.com/Axial-SQL/AxialSqlTools/releases/tag/${args.tag_name}` };
        return state.release;
      }),
      deleteReleaseAsset: record('deleteReleaseAsset', args => {
        state.release.assets = state.release.assets.filter(asset => asset.id !== args.asset_id);
        return {};
      }),
      uploadReleaseAsset: record('uploadReleaseAsset', args => {
        if (state.failUpload === args.name) throw new Error('Upload failed');
        const asset = { id: state.nextAssetId++, name: args.name, state: 'uploaded',
          size: args.data.length, digest: `sha256:${digest(args.data)}` };
        if (state.wrongUploadDigest === args.name) asset.digest = `sha256:${'0'.repeat(64)}`;
        state.release.assets.push(asset);
        return asset;
      }),
      getRelease: record('getRelease', () => state.release || notFound()),
      updateRelease: record('updateRelease', args => {
        state.assetsAtPublication = copy(state.release.assets);
        Object.assign(state.release, args);
        state.latest = state.release;
        return state.release;
      })
    }
  };
  const github = { rest };
  function ownedTag() {
    state.annotation = copy(annotation);
    state.ref = { ref: `refs/tags/${tag}`, object: { type: 'tag', sha: annotation.sha } };
  }
  function ownedRelease({ draft = true, complete = true } = {}) {
    ownedTag();
    state.release = {
      id: 99, tag_name: tag, target_commitish: sha, draft,
      html_url: `https://github.com/Axial-SQL/AxialSqlTools/releases/tag/${tag}`,
      body: `Release notes\n<!-- axial-release-run:${runId};source:${sha} -->`,
      assets: complete ? [vsixName, checksumName].map(name => ({
        id: state.nextAssetId++, name, state: 'uploaded', size: fs.statSync(path.join(repoRoot, 'release-assets', name)).size
      })) : []
    };
    return state.release;
  }
  const methods = () => state.calls.map(call => call.method);
  return { state, context, core, outputs, summaries, env, repoRoot, github,
    ownedTag, ownedRelease, methods,
    prepare: () => prepare({ github, context, core, repoRoot }),
    publish: () => publish({ github, context, core, env, repoRoot }) };
}

test('new release tags the triggering commit, then publishes only after both signed assets upload', async t => {
  const f = fixture(t);
  await f.prepare();
  assert.equal(f.outputs.tag, tag);
  assert.equal(f.outputs.source_sha, sha);
  assert.equal(f.outputs.already_published, 'false');
  const tagCall = f.state.calls.find(call => call.method === 'createTag').args;
  assert.equal(tagCall.object, sha);
  assert.equal(tagCall.type, 'commit');
  const refCall = f.state.calls.find(call => call.method === 'createRef').args;
  assert.equal(refCall.ref, `refs/tags/${tag}`);
  assert.equal(refCall.sha, f.state.annotation.sha);
  await f.publish();
  const creation = f.state.calls.find(call => call.method === 'createRelease').args;
  assert.equal(creation.draft, true);
  assert.equal(creation.make_latest, 'false');
  assert.equal(creation.target_commitish, sha);
  assert.equal(f.state.calls.find(call => call.method === 'generateReleaseNotes').args.previous_tag_name, '4.14');
  assert.deepEqual(f.state.assetsAtPublication.map(asset => asset.name).sort(), [vsixName, checksumName].sort());
  const calls = f.methods();
  assert.ok(calls.lastIndexOf('uploadReleaseAsset') < calls.indexOf('updateRelease'));
  assert.ok(calls.lastIndexOf('getRelease') < calls.indexOf('updateRelease'));
  assert.equal(f.state.release.draft, false);
  assert.equal(f.state.release.make_latest, 'true');
});

test('first release succeeds when there is no existing public release', async t => {
  const f = fixture(t, { latest: null });
  await f.prepare();
  await f.publish();
  assert.equal(f.state.release.draft, false);
  assert.ok(!Object.hasOwn(f.state.calls.find(call => call.method === 'generateReleaseNotes').args, 'previous_tag_name'));
});

for (const previous of ['4.15', 'v4.15.0.0', '4.15.20261002.0930', '4.16']) {
  test(`refuses unchanged or older product version relative to ${previous}`, async t => {
    const f = fixture(t, { latest: { tag_name: previous } });
    await assert.rejects(f.prepare(), /must be newer/);
    assert.ok(!f.methods().includes('createTag'));
    assert.ok(!f.methods().includes('createRelease'));
  });
}

test('fails before tagging when assembly and manifest versions disagree', async t => {
  const f = fixture(t);
  fs.writeFileSync(path.join(f.repoRoot, 'AxialSqlTools/Properties/AssemblyInfo.cs'),
    '[assembly: AssemblyVersion("4.14.0.0")]\n[assembly: AssemblyFileVersion("4.15.0.0")]');
  await assert.rejects(f.prepare(), /AssemblyVersion must be 4\.15\.0\.0/);
  assert.equal(f.state.calls.length, 0);
});

test('requires a manual run from the default branch and matching workflow source SHA', async t => {
  const f = fixture(t);
  f.context.ref = 'refs/heads/feature';
  await assert.rejects(f.prepare(), /must be run manually from main/);
  f.context.ref = 'refs/heads/main';
  f.context.eventName = 'push';
  await assert.rejects(f.prepare(), /must be run manually/);
  f.context.eventName = 'workflow_dispatch';
  f.state.run.head_sha = 'c'.repeat(40);
  await assert.rejects(f.prepare(), /workflow run commit does not match/);
  assert.ok(!f.methods().includes('createTag'));
});

test('a failed build can rerun prepare with the original timestamp and owned tag', async t => {
  const f = fixture(t);
  await f.prepare();
  await f.prepare();
  assert.equal(f.state.calls.filter(call => call.method === 'createTag').length, 1);
  assert.equal(f.state.calls.filter(call => call.method === 'createRef').length, 1);
  assert.equal(f.outputs.tag, tag);
  assert.equal(tagForRun(version, '2026-10-03T15:12:59-05:00'), tag);
});

for (const collision of ['lightweight tag', 'another run', 'another commit']) {
  test(`refuses a minute collision with ${collision}`, async t => {
    const f = fixture(t);
    f.ownedTag();
    if (collision === 'lightweight tag') f.state.ref.object.type = 'commit';
    if (collision === 'another run') f.state.annotation.message = f.state.annotation.message.replace('12345', '99999');
    if (collision === 'another commit') f.state.annotation.object.sha = 'c'.repeat(40);
    await assert.rejects(f.prepare(), /not owned|another run or commit/);
    assert.ok(!f.methods().includes('createRef'));
    assert.ok(!f.methods().includes('createRelease'));
  });
}

test('already published owned release skips rebuild and does not reset latest on retry', async t => {
  const f = fixture(t, { latest: { tag_name: '4.16' } });
  f.ownedRelease({ draft: false });
  await f.prepare();
  assert.equal(f.outputs.already_published, 'true');
  await f.publish();
  assert.equal(f.state.latest.tag_name, '4.16');
  assert.ok(!f.methods().includes('getLatestRelease'));
  for (const mutation of ['createTag', 'createRef', 'createRelease', 'deleteReleaseAsset', 'uploadReleaseAsset', 'updateRelease']) {
    assert.ok(!f.methods().includes(mutation), mutation);
  }
});

test('retry refuses incomplete published release without replacing its assets', async t => {
  const f = fixture(t);
  f.ownedRelease({ draft: false, complete: false });
  await assert.rejects(f.prepare(), /missing a complete|must contain exactly/);
  await assert.rejects(f.publish(), /missing a complete|must contain exactly/);
  assert.ok(!f.methods().includes('uploadReleaseAsset'));
  assert.ok(!f.methods().includes('updateRelease'));
});

for (const change of ['owner', 'source']) {
  test(`refuses an existing release with mismatched ${change}`, async t => {
    const f = fixture(t);
    f.ownedRelease();
    if (change === 'owner') f.state.release.body = 'A manually created draft.';
    if (change === 'source') f.state.release.target_commitish = 'main';
    await assert.rejects(f.prepare(), /was not created by this workflow run/);
    await assert.rejects(f.publish(), /was not created by this workflow run/);
    assert.ok(!f.methods().includes('deleteReleaseAsset'));
    assert.ok(!f.methods().includes('updateRelease'));
  });
}

test('retry resumes an owned draft and replaces its unfinished assets before publishing', async t => {
  const f = fixture(t);
  f.ownedRelease({ complete: false });
  f.state.release.assets = [{ id: 42, name: vsixName, state: 'starter', size: 0 }];
  await f.prepare();
  await f.publish();
  assert.ok(!f.methods().includes('createRelease'));
  assert.equal(f.state.calls.find(call => call.method === 'deleteReleaseAsset').args.asset_id, 42);
  assert.equal(f.state.release.assets.length, 2);
  assert.equal(f.state.release.draft, false);
});

test('an unexpected ZIP asset on an owned draft prevents publication', async t => {
  const f = fixture(t);
  f.ownedRelease();
  f.state.release.assets.push({ id: 42, name: 'unsigned-build.zip', state: 'uploaded', size: 100 });
  await assert.rejects(f.publish(), /unexpected|exactly|only|asset/i);
  assert.equal(f.state.release.draft, true);
  assert.ok(!f.methods().includes('updateRelease'));
  assert.equal(f.state.latest.tag_name, '4.14');
});

for (const problem of ['missing checksum', 'unexpected file', 'wrong version filename', 'changed VSIX', 'wrong signing hash', 'wrong checksum filename']) {
  test(`never creates or publishes a release for ${problem}`, async t => {
    const f = fixture(t);
    f.ownedTag();
    const assets = path.join(f.repoRoot, 'release-assets');
    if (problem === 'missing checksum') fs.unlinkSync(path.join(assets, checksumName));
    if (problem === 'unexpected file') fs.writeFileSync(path.join(assets, 'unsigned.vsix'), 'unexpected');
    if (problem === 'wrong version filename') fs.renameSync(path.join(assets, vsixName), path.join(assets, 'AxialSqlTools_SSMS22_4.14.vsix'));
    if (problem === 'changed VSIX') fs.appendFileSync(path.join(assets, vsixName), 'changed');
    if (problem === 'wrong signing hash') f.env.SIGNED_VSIX_SHA256 = '0'.repeat(64);
    if (problem === 'wrong checksum filename') fs.writeFileSync(path.join(assets, checksumName), `${f.env.SIGNED_VSIX_SHA256}  unsigned.vsix\n`);
    await assert.rejects(f.publish(), /signed artifact must contain exactly|does not match both/);
    assert.ok(!f.methods().includes('createRelease'));
    assert.ok(!f.methods().includes('updateRelease'));
  });
}

test('rejects source or version disagreement between preparation and signed build', async t => {
  const f = fixture(t);
  f.ownedTag();
  f.env.SIGNED_SOURCE_SHA = 'c'.repeat(40);
  await assert.rejects(f.publish(), /same version and commit/);
  f.env.SIGNED_SOURCE_SHA = sha;
  f.env.SIGNED_VERSION = '4.14';
  await assert.rejects(f.publish(), /same version and commit/);
  assert.ok(!f.methods().includes('createRelease'));
});

for (const failure of ['second upload fails', 'server reports wrong digest']) {
  test(`${failure} leaves the release as a draft`, async t => {
    const f = fixture(t);
    f.ownedTag();
    if (failure === 'second upload fails') f.state.failUpload = checksumName;
    else f.state.wrongUploadDigest = vsixName;
    await assert.rejects(f.publish(), /Upload failed|did not pass verification/);
    assert.equal(f.state.release.draft, true);
    assert.ok(!f.methods().includes('updateRelease'));
    assert.equal(f.state.latest.tag_name, '4.14');
  });
}

test('a newer release published while uploading prevents this release taking latest', async t => {
  const f = fixture(t);
  f.ownedTag();
  const upload = f.github.rest.repos.uploadReleaseAsset;
  f.github.rest.repos.uploadReleaseAsset = async args => {
    const result = await upload(args);
    if (args.name === checksumName) f.state.latest = { tag_name: '4.16' };
    return result;
  };
  await assert.rejects(f.publish(), /must be newer/);
  assert.equal(f.state.release.draft, true);
  assert.ok(!f.methods().includes('updateRelease'));
});

test('a GitHub authorization failure is not treated as a missing tag', async t => {
  const f = fixture(t);
  f.github.rest.git.getRef = async () => { throw Object.assign(new Error('Forbidden'), { status: 403 }); };
  await assert.rejects(f.prepare(), /Forbidden/);
  assert.ok(!f.methods().includes('createTag'));
});

test('timestamp tags compare by product version while legacy numeric versions remain supported', () => {
  assert.deepEqual(productVersion('4.15.20261003.2012'), [4, 15, 0, 0]);
  assert.deepEqual(productVersion('v4.15'), [4, 15, 0, 0]);
  assert.deepEqual(productVersion('4.15.2.1'), [4, 15, 2, 1]);
  assert.throws(() => productVersion('4.15.20260230.2012'), /Invalid release timestamp/);
  assert.throws(() => productVersion('4.15.20261003.2460'), /Invalid release timestamp/);
  assert.throws(() => productVersion('4.15.20261003.123'), /Invalid release timestamp/);
});
