# Releasing Axial SQL Tools

The repository has three manually started workflows:

| Workflow | Result |
| --- | --- |
| **Build and Sign VSIX** | Builds and verifies a signed VSIX, then saves it and a SHA256 file as an Actions artifact. It is also reusable by the other workflows. |
| **Prepare New Release** | Creates a tag at the selected source commit, calls Build and Sign VSIX, and publishes a GitHub Release with the signed files. |
| **Publish to galleries** | Reuses an available signed artifact for the selected commit, or calls Build and Sign VSIX, then publishes the same signed VSIX to both galleries. |

Start the two publishing workflows from the repository's default branch, currently `main`. Each run uses the exact commit selected when it starts, even if the branch advances during the build.

## One-time configuration

Keep the existing Azure Artifact Signing configuration in the `vsix-signing` environment. Both callers reuse that environment through Build and Sign VSIX. The Azure login and certificate profile configuration do not change.

For gallery publishing, create these **repository secrets** under **Settings > Secrets and variables > Actions**:

| Secret | Purpose |
| --- | --- |
| `SSMS_GALLERY_MANAGE_TOKEN` | A saved random management password of your choice for the SSMS Gallery listing. |
| `VSIX_GALLERY_MANAGE_TOKEN` | A saved random management password of your choice for the Open VSIX Gallery listing. |

These values are gallery management/delete passwords, not GitHub PATs or API keys issued by the galleries. Choose and save a separate random value for each. The publisher action accepts them through its `manage-token` input.

The action can publish without a token, but on a first upload it prints the gallery's generated management URL, including its token, to the workflow log and summary. This workflow therefore requires the corresponding secret before uploading. An identical package that is already on a gallery can still be skipped without that secret.

The workflows use `GITHUB_TOKEN` for repository operations, including tag/release creation and artifact downloads. No additional GitHub PAT is required. Any existing approval rules on the `vsix-signing` environment still apply.

## Set the product version

Before publishing changed binaries, increment the product's `major.minor` version consistently in these files:

| File | Example for version 4.15 |
| --- | --- |
| `AxialSqlTools/source.extension.vsixmanifest` | `<Identity ... Version="4.15" ... />` |
| `AxialSqlTools/Properties/AssemblyInfo.cs` | `[assembly: AssemblyVersion("4.15.0.0")]` |
| `AxialSqlTools/Properties/AssemblyInfo.cs` | `[assembly: AssemblyFileVersion("4.15.0.0")]` |

Commit the changes and merge them into `main`. The workflows validate that these versions agree; they do not edit or commit a version increment themselves.

GitHub release tags use `major.minor.yyyyMMdd.HHmm`, for example:

```text
4.15.20261003.2012
```

The date and time are **UTC**, taken from the workflow run's original creation time. A rerun keeps the same tag. If another run already owns a tag for that minute, the workflow fails without moving it; start a new run in a later minute.

The date belongs only to the GitHub tag. The VSIX remains `4.15`, and the assembly remains `4.15.0.0`. The updater recognizes the timestamped tag and compares its product version, so it does not offer the same release again after installation. A timestamp does not replace increasing the product version. Prepare New Release rejects a version that is not newer than the latest published GitHub Release, except when retrying its own already completed release.

## Publish a GitHub Release

1. Open **Actions > Prepare New Release > Run workflow**.
2. Select `main` and run it.
3. If the signing environment requires review, complete its existing approval.
4. Wait for **Publish signed release** to succeed. The run summary contains the release URL, source commit, version, and signed VSIX hash.

The workflow creates an annotated tag at the selected commit, verifies that the build checks out that exact commit, and calls Build and Sign VSIX. It downloads the signed artifact by its exact artifact ID and verifies the checksum against the signing job's output.

It then creates a draft release with generated release notes, uploads these two files, verifies their uploads, and publishes the release as **Latest**:

```text
AxialSqlTools_SSMS22_4.15.vsix
AxialSqlTools_SSMS22_4.15.vsix.sha256
```

The Actions artifact ZIP is not a release asset. No published release is exposed until the VSIX and checksum have both uploaded successfully.

## Publish to both galleries

1. Open **Actions > Publish to galleries > Run workflow**.
2. Select `main` and run it.
3. Wait for **Verify and publish to both galleries** to succeed.

For a release distributed everywhere, run Prepare New Release first, then Publish to galleries while `main` still points to the same commit. The gallery workflow will reuse the signed artifact from that release run.

The gallery workflow searches for the exact `AxialSqlTools-Signed-<version>` artifact for the selected SHA. It accepts artifacts from Build and Sign VSIX, Prepare New Release, or a previous Publish to galleries run in this repository. It can reuse a successfully signed artifact even if a later publication step failed. Expired or missing artifacts and artifacts for other commits are not reusable; a new build and signature are required.

Before publishing, the workflow verifies the artifact's file list and SHA256, the VSIX signature, the packaged extension ID/version, and the main assembly's signature and version. The package is not modified after signing. Both galleries receive the same bytes, and the listing README points to that same source commit.

The two destinations are:

- [SSMS Gallery](https://ssmsgallery.azurewebsites.net)
- [Open VSIX Gallery](https://www.vsixgallery.com)

The same `madskristensen/publish-vsixgallery` action publishes to each, with an explicit `gallery-url`. It is pinned to the reviewed commit corresponding to `v1`. These are separate from the Visual Studio Marketplace.

Before each upload, the workflow checks the current gallery listing. It skips an identical signed package, rejects a newer gallery version, and rejects different bytes under the same version. After an upload, it confirms that the gallery serves the expected version and SHA256. The older SSMS Gallery API does not provide SHA256 metadata, so verification downloads and hashes that gallery's VSIX.

## Failures and retries

- **Build or signing failure:** no GitHub Release or gallery upload is published. A tag created by Prepare New Release remains available for the same run's retry.
- **GitHub asset upload failure:** the release remains a draft. Use **Re-run failed jobs**, or rerun the full original workflow, to resume its owned draft. The workflow does not modify a release belonging to another run.
- **Already published GitHub Release:** rerunning its original workflow succeeds without rebuilding or changing its assets. It does not move an older release back to Latest.
- **One gallery fails:** the workflow still attempts the other gallery and reports the overall run as failed. Rerun the failed job; the gallery that already has identical bytes will be skipped.
- **Gallery already has a newer version:** an old run cannot downgrade it. Publish the current intended version instead.
- **Gallery has the same version but different bytes:** increment the product version for a new build. A newly timestamped signature can produce different bytes even from the same source.

Signed Actions artifacts are retained for 30 days. Reuse is possible while the matching artifact is available. If it has expired, gallery publication builds again; do not use that newly signed package to replace an already distributed version. Published GitHub Release assets remain available separately.

## Validate changes to the automation

The release and gallery control logic has tests that use mocked GitHub and gallery APIs:

```bash
node --test .github/scripts/release.test.js .github/scripts/gallery.test.js .github/scripts/gallery-publish.test.js
```

Validate workflow syntax and expressions with `actionlint`. Full VSIX compilation and Authenticode verification run on the Windows runners. Live publication should be exercised by an intentional release, not a test upload under the production extension ID.

Publishing references: [GitHub reusable workflows](https://docs.github.com/en/actions/how-tos/reuse-automations/reuse-workflows), [publisher action](https://github.com/madskristensen/publish-vsixgallery), [SSMS Gallery guide](https://ssmsgallery.azurewebsites.net/devguide), and [Open VSIX Gallery guide](https://www.vsixgallery.com/devguide).
