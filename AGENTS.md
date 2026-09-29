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
   (needs `pip install -r security/requirements.txt`) must print PASS. A
   new olevba or mraptor finding fails it; if the finding is benign, run
   `python security/scan.py --update-baseline`, replace every `TODO` note
   in `security/baseline.json` with the real reason, and review the diff.
   Never accept a finding in the `library` group without reading the code
   that triggers it. "Possible VBA stomping" means the workbook's p-code
   holds names its source lacks; rebuild the VBA project (save a copy as
   .xlsx, re-import every module, save as .xlsm) and compare sheets and
   cells before replacing the workbook. Also run the four
   `pyvbaanalysis` commands from `.github/workflows/vba-analysis.yml`.
6. **Commit, tag, release.** Tag `vx.y.z`, push commit and tag, then
   `gh release create vx.y.z` with BOTH `dist/*.bas` files attached as
   assets. Confirm new test files actually appear in the staged list
   (`A Tests/...`) before pushing. Publishing the release triggers
   `.github/workflows/release-security-report.yml`, which rescans the tag,
   checks the report's hashes against the release's `.bas` files, and
   attaches `security-report.md` and `security-report.json`; check they
   appear on the release.

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
