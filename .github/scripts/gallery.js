'use strict';

const { readVersion } = require('./release.js');

const SIGNING_WORKFLOWS = new Set([
  '.github/workflows/build-sign-vsix.yml',
  '.github/workflows/prepare-new-release.yml',
  '.github/workflows/publish-to-galleries.yml'
]);

async function findSignedArtifact({ github, context, core, repoRoot = process.cwd() }) {
  const repository = context.payload.repository;
  const defaultBranch = repository?.default_branch;
  if (!defaultBranch || context.eventName !== 'workflow_dispatch' || context.ref !== `refs/heads/${defaultBranch}`) {
    throw new Error(`Publish to galleries must be run manually from ${defaultBranch || 'the repository default branch'}.`);
  }

  const version = readVersion(repoRoot);
  const artifactName = `AxialSqlTools-Signed-${version}`;
  const repositoryName = `${context.repo.owner}/${context.repo.repo}`.toLowerCase();
  const matchesId = id => id == null || repository.id == null || String(id) === String(repository.id);
  const matchesRepository = candidate => !candidate ||
    (matchesId(candidate.id) && (!candidate.full_name || candidate.full_name.toLowerCase() === repositoryName));
  const candidates = [];

  // The name filter avoids scanning unrelated artifacts. Read every page before
  // sorting because the repository endpoint does not promise an ordering.
  for (let page = 1; ; page++) {
    const { data } = await github.rest.actions.listArtifactsForRepo({
      ...context.repo, name: artifactName, per_page: 100, page
    });
    for (const artifact of data.artifacts) {
      if (artifact.name === artifactName && !artifact.expired && artifact.size_in_bytes > 0 &&
          artifact.workflow_run?.head_sha === context.sha &&
          matchesId(artifact.workflow_run.repository_id) && matchesId(artifact.workflow_run.head_repository_id)) {
        candidates.push(artifact);
      }
    }
    if (data.artifacts.length < 100) break;
  }

  let selected;
  let selectedRun;
  const runs = new Map();
  for (const artifact of candidates.sort((left, right) => right.id - left.id)) {
    const runId = artifact.workflow_run.id;
    if (!runs.has(runId)) {
      try {
        const { data } = await github.rest.actions.getWorkflowRun({ ...context.repo, run_id: runId });
        runs.set(runId, data);
      } catch (error) {
        if (error.status !== 404) throw error;
        // A run may have been deleted after the artifact listing was returned.
        runs.set(runId, null);
      }
    }
    const run = runs.get(runId);
    // GitHub can include the workflow ref after the path, for example @main.
    const workflowPath = run?.path?.split('@', 1)[0];
    if (!run || run.head_sha !== context.sha || run.event !== 'workflow_dispatch' ||
        !SIGNING_WORKFLOWS.has(workflowPath) ||
        !matchesRepository(run.repository) || !matchesRepository(run.head_repository)) continue;

    // Build and Sign VSIX uploads this artifact only after signing and signature
    // verification. A later release/gallery upload may fail, and a retry may be
    // in progress, so the overall run does not need a successful conclusion.
    // The exact commit matters; a manual build of its tag is also reusable.
    selected = artifact;
    selectedRun = run;
    break;
  }

  const outputs = {
    version,
    source_sha: context.sha,
    signed_artifact_id: selected ? String(selected.id) : '',
    signed_run_id: selected ? String(selected.workflow_run.id) : ''
  };
  for (const [name, value] of Object.entries(outputs)) core.setOutput(name, value);

  const runUrl = selectedRun?.html_url ||
    `${context.serverUrl || 'https://github.com'}/${context.repo.owner}/${context.repo.repo}/actions/runs/${outputs.signed_run_id}`;
  await core.summary.addRaw(selected
    ? `## Signed package ready\n\nVersion: ${version}\n\nSource commit: ${context.sha}\n\nReuse artifact ${selected.id} from [workflow run ${outputs.signed_run_id}](${runUrl}).\n`
    : `## Signed package required\n\nVersion: ${version}\n\nSource commit: ${context.sha}\n\nNo reusable signed artifact was found for this exact commit. Build and Sign VSIX will prepare one.\n`
  ).write();
  return outputs;
}

module.exports = { findSignedArtifact };
