# Releasing Axial SQL Tools

Run the publishing workflows manually from `main`. Each run uses the commit selected when it starts.

| Workflow | Result |
| --- | --- |
| **Build and Sign VSIX** | Builds, signs, and verifies the VSIX, then saves it and its SHA256 checksum as an Actions artifact. |
| **Prepare New Release** | Checks the version tag, builds and signs, creates the tag, and publishes a GitHub prerelease with the signed files. |
| **Publish to galleries** | Reuses a signed artifact for the same commit if available, otherwise builds and signs, then uploads to both galleries. |

## Prepare a release

1. Update `AxialSqlTools/source.extension.vsixmanifest` and the assembly versions in `AxialSqlTools/Properties/AssemblyInfo.cs`. For example, use `4.15` in the manifest and `4.15.0.0` for both assembly versions.
2. Commit and merge into `main`.
3. Run **Actions > Prepare New Release** from `main`.
4. Review the prerelease. When ready, edit it on GitHub, clear **Set as a pre-release**, and select **Set as the latest release**.

The tag is exactly the manifest version, for example `4.15`. There is no date or time suffix. The workflow fails if that tag already exists, including on a rerun after successful publication. It never replaces an existing release or automatically makes one Latest.

The tag is created at the selected commit after building and signing succeed. The release action stages the signed VSIX and checksum in a draft before publishing the prerelease. Assets are named `AxialSqlTools_SSMS22_4.15.vsix` and `AxialSqlTools_SSMS22_4.15.vsix.sha256`.

If building or signing fails, retry the run. If publication fails after tag creation, inspect the tag and any draft release before retrying. To retry the same version, manually remove only that failed attempt's unpublished tag/draft; otherwise increment the version. There is no automatic cleanup or overwrite.

## Publish to both galleries

Run **Actions > Publish to galleries** from `main`. To reuse the exact signed package from a release, run it while `main` still points to the same commit. Signed artifacts are retained for 30 days. The workflow looks for an unexpired signed artifact from one of the three workflows above, including runs where signing succeeded but later publication failed.

Each gallery upload uses the documented [`madskristensen/publish-vsixgallery@v1`](https://github.com/madskristensen/publish-vsixgallery) action:

- [SSMS Gallery](https://ssmsgallery.azurewebsites.net/devguide)
- [Open VSIX Gallery](https://www.vsixgallery.com/devguide)

Configure repository secrets `SSMS_GALLERY_MANAGE_TOKEN` and `VSIX_GALLERY_MANAGE_TOKEN` with saved management passwords of your choice. These are gallery management passwords, not GitHub PATs. The action accepts them through `manage-token`; without one, a first upload may print an automatically generated management URL in the public run log.

The gallery workflow publishes directly and stops if an upload fails. It does not add custom gallery version checks or retry logic. Increment the product version before publishing changed binaries. GitHub prerelease status does not delay gallery publication; run the gallery workflow when the package is ready for gallery users.

The existing Azure configuration and approval rules in the `vsix-signing` environment also apply to the reusable build workflow. No additional GitHub PAT is required.
