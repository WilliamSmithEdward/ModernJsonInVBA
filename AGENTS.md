# Working in this repo

Guidance for AI agents (and humans) making changes to ModernJsonInVBA.

## Release workflow

Follow this order; the version stamps depend on it.

1. **CHANGELOG.md first.** Add the release entry at the top:
   `## [x.y.z] - YYYY-MM-DD`, plus the link line at the bottom of the file.
   This entry is the single source of the version and date; nothing else is
   edited by hand for a version bump.
2. **`python build_dist.py`.** Stamps `Version:` / `Released:` into every
   `vba_source/*.bas` module header and both dist file headers from the top
   CHANGELOG entry, rebuilds `dist/`, and runs the portability check (the
   AllO365 build must contain no Excel references). Idempotent: re-running
   without a version change touches nothing.
3. **Sync the workbook and run every test suite.** Import the refreshed
   `vba_source/*.bas` modules into `ModernJsonInVBA.xlsm` via Excel COM
   (`VBComponents.Import`; pyOpenVBA can read and replace module source but
   cannot add modules). Run all `RunAll_*` / `Json_RunAllTests` macros on
   a patched copy whose test-module `MsgBox` lines are replaced with
   `Debug.Print` first. Drive the runs with pyvbaharness
   (`pip install pyvbaharness`): a failing assert raised through a suite
   runner otherwise becomes a modal dialog that hangs headless
   automation, while the harness returns it as data and kills its own
   Excel on timeout. Its run targets are capped at 31 characters, so call
   each runner through a short wrapper procedure.
4. **Anti-smell scan.** All comments and docs are pure ASCII except
   functional arrows in diagrams: no em/en dashes, smart quotes, ellipsis
   character, or multiplication sign, and no unsupported frequency claims
   ("APIs usually...") in prose. Test-data unicode inside string literals is
   intentional; never "fix" it. Scan with Python, not grep.
5. **Security scan and static analysis.** `python security/scan.py`
   (needs `pip install --require-hashes -r
   .github/requirements/security.txt`) must print PASS. A
   new olevba or mraptor finding fails it; if the finding is benign, run
   `python security/scan.py --update-baseline`, replace every `TODO` note
   in `security/baseline.json` with the real reason, and review the diff.
   Never accept a finding in the `library` group without reading the code
   that triggers it. "Possible VBA stomping" means the workbook's p-code
   holds names its source lacks; rebuild the VBA project (save a copy as
   .xlsx, re-import every module, save as .xlsm) and compare sheets and
   cells before replacing the workbook. Also run the four
   `pyvbaanalysis` commands from `.github/workflows/ci.yml`.
6. **Merge, dry-run, tag.** Commit the changelog, the stamped sources, the
   rebuilt `dist/` and the synced workbook, and confirm new test files
   actually appear in the staged list (`A Tests/...`). Merge the pull
   request to `main`. Optionally dry-run the release:
   `gh workflow run publish.yml --ref main`, then
   `gh run download <run-id> -n release-preview`. The owner then tags the
   merged commit: `git tag vx.y.z && git push origin vx.y.z`. The tag runs
   `.github/workflows/publish.yml`: it refuses a tag that is not the top
   CHANGELOG version, reruns `build_dist.py` and fails if the committed
   `dist/` differs, runs Security and Malware scan, checks the combined
   report's hashes against the files it built, signs their build
   provenance, and creates the release with both `dist/*.bas` files,
   `ModernJsonInVBA-x.y.z.sigstore.json`, `security-report.md` and
   `security-report.json`, with the CHANGELOG section as its notes. The
   workbook is not attached. Do not create the release by hand.

## Project constraints

- Pure VBA: no `Scripting.Dictionary`, no COM references, no `LongLong`
  (32-bit Office), no `Declare` statements (Mac hosts).
- The public API and error numbers are frozen; see the README's
  Deterministic Errors section before changing any raise.
- VBA requires every module-level `Const` / `Type` / `Enum` to precede all
  procedures in the module.
- New features ship as minor versions, fixes as patches, and each release
  gets a CHANGELOG entry in Keep a Changelog format.
- Performance claims in README / PERFORMANCE.md regenerate from
  `Run_JsonPerfMatrix` (payloads from `json_payloads/generate_payloads.py`);
  update the numbers by re-running, not by editing.
- Conformance claims trace to CONFORMANCE.md (JSONTestSuite); re-run the
  corpus when the parser changes.

## Layout

- `vba_source/` - the twelve library modules (source of truth)
- `dist/` - generated single-file builds; never edit by hand
- `Tests/` - repo test suites, imported into the workbook beside the
  workbook-only legacy suites (`Tests_JsonParser_` and friends)
- `ModernJsonInVBA.xlsm` - the shipping workbook with all modules and tests
- `security/` - olevba/mraptor scan script and its reviewed baseline; see
  SECURITY.md

<!-- repo-standards:begin. Copied from WilliamSmithEdward/repo-standards, templates/agents/AGENTS-block.md. Change it there; the weekly rescan fails a copy that differs. -->
## Releases, CI and security

These rules are the same in every WilliamSmithEdward repository.

- **How a release happens here:** pushing a `vX.Y.Z` tag runs Publish, which builds the release files in CI and creates the GitHub release with them, their signed provenance and the security reports. Any other step, such as a marketplace upload, is described elsewhere in this file.
- **Starting a workflow by hand never releases anything.** Publish and every
  release report are dry runs when started with `gh workflow run` or the Run
  workflow button. They build, scan and assemble the release files exactly
  as a release would, and upload them as the `release-preview` artifact
  instead. Run one after changing anything on the release path:
  `gh workflow run <file> --ref main`, then
  `gh run download <run-id> -n release-preview`.
- **Do not create, publish, edit or delete a release or a `v*` tag** unless
  the owner asks for it. A `v*` tag cannot be moved or deleted once pushed.
- **Every change to `main` goes through a pull request** that passes CI
  passed, Security passed and Malware scan passed. No one can push to `main`
  directly or skip the checks, admins included. Push a branch, open a pull
  request, and let it merge itself: `gh pr merge --auto --squash <number>`.
- **Pins.** Actions by full commit SHA with the version as a comment. Images
  by digest, in `.github/security/<tool>/Dockerfile`. Python tools from the
  hash-locked `.github/requirements/<purpose>.txt`, compiled from the `.in`
  beside it with
  `uv pip compile <purpose>.in --universal --generate-hashes --python-version 3.12 -o <purpose>.txt`.
  Runners are named releases, never `-latest`.
- **Updates merge themselves.** Dependabot and the Update YARA rules workflow
  open pull requests that merge once the three checks pass, except a
  third-party major version, which waits for the owner. Leave them alone
  unless asked.
- **A scanner finding is fixed or accepted with a written reason** in the
  repository's accepted list. Never silence a scanner without one.
<!-- repo-standards:end -->
