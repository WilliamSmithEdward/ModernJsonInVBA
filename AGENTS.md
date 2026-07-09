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
   cannot add modules). Run all `RunAll_*` / `Json_RunAllTests` macros
   headlessly on a patched copy whose test-module `MsgBox` lines are
   replaced with `Debug.Print` first. A failing assert raised through a
   suite runner becomes a modal dialog that hangs headless automation, so
   drive runs with a timeout and be ready to kill EXCEL.EXE and diagnose
   per-test.
4. **Anti-smell scan.** All comments and docs are pure ASCII except
   functional arrows in diagrams: no em/en dashes, smart quotes, ellipsis
   character, or multiplication sign, and no unsupported frequency claims
   ("APIs usually...") in prose. Test-data unicode inside string literals is
   intentional; never "fix" it. Scan with Python, not grep.
5. **Commit, tag, release.** Tag `vx.y.z`, push commit and tag, then
   `gh release create vx.y.z` with BOTH `dist/*.bas` files attached as
   assets. Confirm new test files actually appear in the staged list
   (`A Tests/...`) before pushing.

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
