'use strict';

const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

function productVersion(tag) {
  const match = /^v?(\d+)\.(\d+)(?:\.(\d+))?(?:\.(\d+))?$/i.exec(tag || '');
  if (!match) throw new Error(`Unsupported release version: ${tag}`);
  const parts = match.slice(1).map(value => value === undefined ? 0 : Number(value));
  if (parts.some(value => !Number.isSafeInteger(value) || value > 2147483647)) {
    throw new Error(`Invalid release version: ${tag}`);
  }
  // Keep this timestamp convention aligned with UpdateChecker.ParseVersion.
  if (match[3]?.length === 8) {
    const date = match[3];
    const time = match[4];
    if (!time || time.length !== 4) throw new Error(`Invalid release timestamp: ${tag}`);
    const iso = `${date.slice(0, 4)}-${date.slice(4, 6)}-${date.slice(6, 8)}T${time.slice(0, 2)}:${time.slice(2, 4)}:00.000Z`;
    const parsed = new Date(iso);
    if (!Number(date.slice(0, 4)) || Number.isNaN(parsed.valueOf()) || parsed.toISOString() !== iso) {
      throw new Error(`Invalid release timestamp: ${tag}`);
    }
    parts[2] = 0;
    parts[3] = 0;
  }
  return parts;
}

function readVersion(repoRoot) {
  const manifest = fs.readFileSync(path.join(repoRoot, 'AxialSqlTools/source.extension.vsixmanifest'), 'utf8');
  const identity = /<Identity\b[^>]*>/.exec(manifest)?.[0];
  const version = /\bVersion\s*=\s*['"]([^'"]+)['"]/.exec(identity || '')?.[1];
  if (!/^(0|[1-9]\d*)\.(0|[1-9]\d*)$/.test(version || '') ||
      version.split('.').some(part => Number(part) > 65534)) {
    throw new Error('The release manifest must use a major.minor version, for example 4.15.');
  }
  const assemblyInfo = fs.readFileSync(path.join(repoRoot, 'AxialSqlTools/Properties/AssemblyInfo.cs'), 'utf8');
  for (const attribute of ['AssemblyVersion', 'AssemblyFileVersion']) {
    const pattern = new RegExp(`^\\s*\\[assembly:\\s*${attribute}\\("([^"]+)"\\)\\]`, 'm');
    if (pattern.exec(assemblyInfo)?.[1] !== `${version}.0.0`) {
      throw new Error(`${attribute} must be ${version}.0.0 to match the VSIX manifest.`);
    }
  }
  return version;
}

function tagForRun(version, createdAt) {
  const created = new Date(createdAt);
  if (Number.isNaN(created.valueOf())) throw new Error('The workflow run creation time is invalid.');
  const iso = created.toISOString();
  return `${version}.${iso.slice(0, 10).replaceAll('-', '')}.${iso.slice(11, 16).replace(':', '')}`;
}

async function optional(request) {
  try { return (await request()).data; }
  catch (error) {
    if (error.status === 404) return null;
    throw error;
  }
}

function ownership(context, sha) {
  return `<!-- axial-release-run:${context.runId};source:${sha} -->`;
}

function tagMessage(context, sha, version) {
  return `Prepare New Release\nRun: ${context.runId}\nSource: ${sha}\nVersion: ${version}`;
}

async function verifyTag(github, context, tag, sha, version) {
  const ref = await optional(() => github.rest.git.getRef({ ...context.repo, ref: `tags/${tag}` }));
  if (!ref) return false;
  if (ref.object.type !== 'tag') throw new Error(`Tag ${tag} already exists and is not owned by this workflow run.`);
  const { data: annotation } = await github.rest.git.getTag({ ...context.repo, tag_sha: ref.object.sha });
  if (annotation.tag !== tag || annotation.object.type !== 'commit' || annotation.object.sha !== sha ||
      annotation.message.trim() !== tagMessage(context, sha, version)) {
    throw new Error(`Tag ${tag} belongs to another run or commit. Start a new run in a later UTC minute.`);
  }
  return true;
}

function assertOwnedRelease(release, context, sha) {
  if (!release.body?.includes(ownership(context, sha)) || release.target_commitish !== sha) {
    throw new Error(`Release ${release.tag_name} was not created by this workflow run; it will not be changed.`);
  }
}

function assertCompleteRelease(release, version) {
  const names = [`AxialSqlTools_SSMS22_${version}.vsix`, `AxialSqlTools_SSMS22_${version}.vsix.sha256`];
  if (JSON.stringify((release.assets || []).map(asset => asset.name).sort()) !== JSON.stringify([...names].sort())) {
    throw new Error(`Release ${release.tag_name} must contain exactly the signed VSIX and checksum, with no unexpected assets.`);
  }
  for (const name of names) {
    const assets = (release.assets || []).filter(asset => asset.name === name && asset.state === 'uploaded' && asset.size > 0);
    if (assets.length !== 1) throw new Error(`Published release ${release.tag_name} is missing a complete ${name} asset.`);
  }
}

async function assertNewVersion(github, context, version) {
  const latest = await optional(() => github.rest.repos.getLatestRelease(context.repo));
  if (latest) {
    const current = productVersion(version);
    const previous = productVersion(latest.tag_name);
    const difference = current.map((value, index) => value - previous[index]).find(value => value !== 0) || 0;
    if (difference <= 0) {
      throw new Error(`Version ${version} must be newer than published release ${latest.tag_name}. Increase the manifest and assembly versions before starting a new release.`);
    }
  }
  return latest;
}

async function summary(core, text) {
  await core.summary.addRaw(text).write();
}

async function prepare({ github, context, core, repoRoot = process.cwd() }) {
  const defaultBranch = context.payload.repository.default_branch;
  if (context.eventName !== 'workflow_dispatch' || context.ref !== `refs/heads/${defaultBranch}`) {
    throw new Error(`Prepare New Release must be run manually from ${defaultBranch}.`);
  }
  const version = readVersion(repoRoot);
  const { data: run } = await github.rest.actions.getWorkflowRun({ ...context.repo, run_id: context.runId });
  if (run.head_sha !== context.sha) throw new Error('The workflow run commit does not match the checked-out release source.');
  const tag = tagForRun(version, run.created_at);
  const exists = await verifyTag(github, context, tag, context.sha, version);
  const release = await optional(() => github.rest.repos.getReleaseByTag({ ...context.repo, tag }));

  if (release) {
    if (!exists) throw new Error(`Release ${tag} has no matching tag owned by this run.`);
    assertOwnedRelease(release, context, context.sha);
  }
  const published = Boolean(release && !release.draft);
  if (published) {
    assertCompleteRelease(release, version);
  } else {
    await assertNewVersion(github, context, version);
    if (!exists) {
      const { data: annotation } = await github.rest.git.createTag({
        ...context.repo, tag, message: tagMessage(context, context.sha, version), object: context.sha, type: 'commit'
      });
      await github.rest.git.createRef({ ...context.repo, ref: `refs/tags/${tag}`, sha: annotation.sha });
    }
  }

  core.setOutput('version', version);
  core.setOutput('tag', tag);
  core.setOutput('source_sha', context.sha);
  core.setOutput('already_published', String(published));
  await summary(core, published
    ? `## Release already published\n\n[${tag}](${release.html_url}) is complete. This retry leaves it unchanged.\n`
    : `## Release prepared\n\nVersion: ${version}\n\nTag: ${tag} (UTC)\n\nSource commit: ${context.sha}\n\nThe signed package will be published after both signature checks succeed.\n`);
}

function readAssets(directory, version, expectedHash) {
  const vsixName = `AxialSqlTools_SSMS22_${version}.vsix`;
  const names = [vsixName, `${vsixName}.sha256`];
  const found = fs.readdirSync(directory).sort();
  if (JSON.stringify(found) !== JSON.stringify([...names].sort())) {
    throw new Error('The signed artifact must contain exactly the expected VSIX and its SHA256 file.');
  }
  const files = names.map(name => ({ name, data: fs.readFileSync(path.join(directory, name)) }));
  const hash = crypto.createHash('sha256').update(files[0].data).digest('hex');
  const checksum = /^([a-f0-9]{64})  ([^\r\n]+)\r?\n?$/i.exec(files[1].data.toString('ascii'));
  if (!files[0].data.length || !/^[a-f0-9]{64}$/i.test(expectedHash || '') || hash !== expectedHash.toLowerCase() ||
      !checksum || checksum[1].toLowerCase() !== hash || checksum[2] !== vsixName) {
    throw new Error('The downloaded VSIX does not match both the signing job hash and its checksum file.');
  }
  return files;
}

async function publish({ github, context, core, env, repoRoot = process.cwd() }) {
  const version = readVersion(repoRoot);
  const tag = env.RELEASE_TAG;
  const sha = env.RELEASE_SOURCE_SHA;
  if (env.RELEASE_VERSION !== version || env.SIGNED_VERSION !== version ||
      sha !== context.sha || env.SIGNED_SOURCE_SHA !== sha ||
      !tag?.startsWith(`${version}.`) || productVersion(tag).join('.') !== `${version}.0.0`) {
    throw new Error('The prepared release and signed build do not identify the same version and commit.');
  }
  if (!await verifyTag(github, context, tag, sha, version)) throw new Error(`The prepared tag ${tag} is missing.`);
  const files = readAssets(path.join(repoRoot, 'release-assets'), version, env.SIGNED_VSIX_SHA256);
  let release = await optional(() => github.rest.repos.getReleaseByTag({ ...context.repo, tag }));
  if (release) {
    assertOwnedRelease(release, context, sha);
    if (!release.draft) {
      assertCompleteRelease(release, version);
      await summary(core, `## Release already published\n\n[${tag}](${release.html_url}) is complete. This retry leaves it unchanged.\n`);
      return;
    }
  }
  const previous = await assertNewVersion(github, context, version);

  if (!release) {
    const { data: notes } = await github.rest.repos.generateReleaseNotes({
      ...context.repo, tag_name: tag, target_commitish: sha,
      ...(previous ? { previous_tag_name: previous.tag_name } : {})
    });
    const runUrl = `${context.serverUrl}/${context.repo.owner}/${context.repo.repo}/actions/runs/${context.runId}`;
    const body = `${notes.body}\n\n## Build\n\n- Extension version: ${version}\n- Source commit: ${sha}\n- [Build and signature verification](${runUrl})\n\n${ownership(context, sha)}`;
    ({ data: release } = await github.rest.repos.createRelease({
      ...context.repo, tag_name: tag, target_commitish: sha,
      name: `Axial SQL Tools ${version} | SSMS 22 VSIX`, body,
      draft: true, prerelease: false, make_latest: 'false'
    }));
  }

  // Only a draft owned by this run can have unfinished assets replaced on retry.
  for (const file of files) {
    for (const existing of (release.assets || []).filter(asset => asset.name === file.name)) {
      await github.rest.repos.deleteReleaseAsset({ ...context.repo, asset_id: existing.id });
    }
    const { data: asset } = await github.rest.repos.uploadReleaseAsset({
      ...context.repo, release_id: release.id, name: file.name, data: file.data,
      headers: { 'content-type': 'application/octet-stream', 'content-length': file.data.length }
    });
    const digest = `sha256:${crypto.createHash('sha256').update(file.data).digest('hex')}`;
    if (asset.name !== file.name || asset.state !== 'uploaded' || asset.size !== file.data.length ||
        (asset.digest && asset.digest.toLowerCase() !== digest)) {
      throw new Error(`Uploaded asset ${file.name} did not pass verification; the release remains a draft.`);
    }
  }

  // Another release might have been published manually while this run was building.
  await assertNewVersion(github, context, version);
  const { data: ready } = await github.rest.repos.getRelease({ ...context.repo, release_id: release.id });
  assertOwnedRelease(ready, context, sha);
  assertCompleteRelease(ready, version);
  if (!ready.draft) throw new Error('The release was published outside this job while its assets were uploading.');
  const { data: published } = await github.rest.repos.updateRelease({
    ...context.repo, release_id: release.id, draft: false, prerelease: false, make_latest: 'true'
  });
  await summary(core, `## Release published\n\n[${tag}](${published.html_url})\n\nVersion: ${version}\n\nSource commit: ${sha}\n\nSigned VSIX SHA256: ${env.SIGNED_VSIX_SHA256}\n`);
}

module.exports = { prepare, publish, productVersion, readVersion, tagForRun, readAssets };
