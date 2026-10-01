# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/ModernJsonInVBA/security/advisories/new).
Only the maintainer sees it. Include the release you used and the host
application (Excel, Word, Access, PowerPoint), and the smallest payload or
macro that shows it, with credentials and private data removed.

A confirmed vulnerability is fixed in a release on the GitHub releases
page, and the advisory is published with it,
crediting you unless you ask otherwise.

## Supported versions

Only the latest release on the GitHub releases page receives security
fixes. Older releases are not maintained separately; update when a fix
ships.

## Scope

The library is the twelve modules in `vba_source/` and the two single-file
builds in `dist/` that `build_dist.py` generates from them. The code is
plain VBA:

- No `Declare` statements, so no Windows or Mac API calls.
- No `CreateObject` or `GetObject`, so no COM objects, no scripting host,
  and no network access. The HTTP helpers in the README are examples for
  your own code; they are not part of the library.
- No `Shell`, no auto-run procedures (`AutoOpen`, `Workbook_Open`, and so
  on), and no reads of environment variables.
- One file operation: `Json_ReadTextFile` (also used by the CSV and NDJSON
  file readers) opens the path you pass it `For Binary Access Read`.
  Nothing in the library writes, deletes, or lists files.
- The Excel modules write only to the worksheet and table you pass in.

A vulnerability here is library code that does more than this: reads a
file it was not passed, writes outside the worksheet and table it was
given, or evaluates payload content as a formula when you asked it not to.

### Loading JSON you do not trust

By default, a JSON string value that begins with `=` is written into the
cell as a formula, the same as if you had typed it. The payload
`[{"f":"=1+1"}]` produces a cell that holds the formula `=1+1` and shows
`2`. This is deliberate: it is how formulas travel through the library,
for example with `preserveFormulas` exports and coalesce literals. A key
that begins with `=` is evaluated once when it becomes a column header.
Strings that begin with `+`, `-`, or `@` stay text.

If the payload comes from a source you do not control, a value such as
`=WEBSERVICE(...)` or `=HYPERLINK(...)` would become a live formula in your
workbook. Pass `formulaStringsAsText:=True` to the `Excel_Upsert...`
function you call:

```vba
Excel_UpsertListObjectFromJsonAtRoot ws, "tblOrders", ws.Range("A1"), jsonText, _
    "$.data", formulaStringsAsText:=True
```

Every value and header that begins with `=` or `'` is then written with an
apostrophe prefix, so the cell holds the exact text and nothing from the
payload is evaluated. Numbers, booleans, and other text are unchanged, and
formula columns you added to the table yourself are still preserved. This
was checked in Excel for Microsoft 365 on Windows and is pinned by the
`Tests_FormulaText` suite.

### The example workbook

`ModernJsonInVBA.xlsm` contains the library plus test suites and examples.
Some of those macros download sample JSON from public demo APIs
(pokeapi.co, dummyjson.com, jsonplaceholder.typicode.com) with
`MSXML2.XMLHTTP`, and some write and delete temporary files in `%TEMP%`.
None of them run automatically. If you only need the library, import a
file from `dist/` instead of using the workbook.

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security and Malware scan also run daily at 08:17 UTC, and again
on the tagged commit when a release is published.

- **Code:** olevba and mraptor, from oletools, scan every VBA file in three
  groups: the library (`vba_source/`, `dist/`), the tests (`Tests/`,
  `json_payloads/`) and the workbook. A tracked VBA file outside every
  group fails the scan. olevba flags the keywords malicious macros depend
  on: file and process access, COM objects, auto-run entry points, encoded
  strings, and URLs. mraptor looks for the combination malicious macros
  need: an auto-run trigger plus a way to write files or execute code.
  `security/scan.py` runs both and writes the results to the run's
  `security-report` artifact, not to code scanning. CI also runs
  pyVBAanalysis over the sources, both builds and the workbook for
  compile, type, and dead-code errors; it checks correctness, not intent.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** there is nothing to audit. The library has no
  dependencies, and the Python tools the workflows run come from hash-locked
  files (see Pinning and updates).
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan the twelve source modules, both generated builds, the
  test and payload modules, and the workbook. ClamAV also scans the VBA
  extracted from those files. YARA-X also runs local rules
  (`security/vba_malware.yar`) over that source for encoded PowerShell,
  remote execution through Windows binaries, and Office Run key
  persistence. A YARA Forge download that does not match its pinned
  SHA-256, an unavailable scanner, a failed signature update, or a rule
  compilation error fails Malware scan.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  Some of its checks do not fit this project. A single maintainer cannot
  have a second person approve every change. The release files are built
  locally rather than by CI, so they carry the scan report's SHA-256 list
  rather than a build provenance signature. Fuzzing does not apply: the
  parser is VBA, which runs only inside Office. The JSONTestSuite
  conformance run (CONFORMANCE.md) exercises the parser on malformed input
  instead.

### The VBA stomping check

olevba raises a "VBA Stomping" flag on the workbook. Stomping means the
compiled p-code Excel runs differs from the source you can read. olevba
compares the names and strings in the p-code with the source and stops at
the first one missing. On this workbook the misses are type suffixes that
the p-code dump adds (`esc$` for a name the source spells `esc`). That flag
is never accepted. Instead, `security/scan.py` repeats the comparison for
every name in the p-code and fails the scan on any miss other than such a
suffix. Identifiers match without regard to case, as VBA treats them.

That check found real residue in the workbook shipped with 3.8.2: three
names from deleted code (`Module1`, `PP_SortNames`, `tableRootLabel`) in
the p-code. They were leftovers from earlier edits, not hidden code, and
the workbook's VBA project has since been rebuilt from its module sources.

### Limits of the scan

The scanners inspect static content. They do not run the code, and a clean
report does not prove the code is safe. Reading the source is the stronger
check, and the library is short enough to read.

### Running the scans locally

```bash
python -m pip install --require-hashes -r .github/requirements/security.txt
python security/scan.py
```

For the malware scan as well, install ClamAV, update its signatures with
`freshclam`, install `.github/requirements/malware.txt` the same way, then
run (set `CLAMAV_DATABASE` if its database is in a nondefault directory):

```bash
python security/malware_scan.py --out security-report/malware-results.json
python security/scan.py --malware-results security-report/malware-results.json
```

## Accepted findings

A finding is fixed, or accepted with a written reason in
`security/baseline.json` (olevba and mraptor) or
`security/malware-exceptions.json` (ClamAV and YARA-X).

A baseline entry matches a finding's kind and keyword within one group, and
every entry has a note saying why it is there. Each group has its own list,
so a finding accepted for the tests is still a failure in the library. Each
group also lists the URL hosts its files may name; any other host fails the
scan. Baseline entries are not tied to a file's content, and an entry that
is no longer seen is reported, not failed.

A malware exception matches the exact scanner, signature, file path and
file SHA-256, so a changed file needs another review, and an exception
that no longer matches fails the scan. Do not bypass a scanner failure or
allow an entire rule collection.

zizmor keeps its exceptions in `.github/zizmor.yml` or inline beside the
line they excuse, each with its reason.

Current entries:

- zizmor: one rule is turned off in `.github/zizmor.yml`,
  `self-repository`, which asks for GitHub's `$/` syntax in `uses:`. It
  comes back once GitHub's documentation confirms that syntax for called
  reusable workflows. There are no inline exceptions.
- Baseline: 20 findings in the library, 24 in the tests, and 35 in the
  workbook. Most are ordinary words in code and comments, such as "open"
  in "open-addressing". Allowed hosts are github.com for the library,
  pokeapi.co for the tests, and pokeapi.co, dummyjson.com and
  jsonplaceholder.typicode.com for the workbook.
- mraptor rates the Excel build and the workbook SUSPICIOUS on three false
  matches, all explained in the baseline: the public procedure
  `Excel_ResizeTableToRowCol` fits its pattern for event handlers like
  `UserForm_Resize`, it counts the read-only
  `Open ... For Binary Access Read` in `Json_ReadTextFile` as a write, and
  it reads "run" in comments as an execute. The per-file verdict is shown
  in the report; each match is checked against the baseline.
- Malware exceptions: there are none.

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, the ClamAV image to a digest, Python tools
(oletools, YARA-X, zizmor, pyVBAanalysis) to hash-locked lock files in
`.github/requirements/`, and the YARA Forge rules to a release and its
SHA-256 in `.github/security/yara.json`. ClamAV's signatures change too
often to pin, so freshclam fetches and verifies them on every run.

Dependabot proposes updates to GitHub Actions, the Python lock files and
the ClamAV image once a version is a week old (the owner's own packages,
such as pyVBAanalysis, at once), and at once for a security advisory. The
Update YARA rules workflow proposes new YARA pins each week. A minor or
patch update, and the YARA pull request, merges itself once CI, Security
and Malware scan pass; a third-party major version waits for review.

## Releases

A release is built locally: `build_dist.py` stamps the version and
generates `dist/` from `vba_source/`, and the GitHub release carries the
two `.bas` builds. The workbook is not a release file; get it from the
repository at the release tag.

Publishing the release starts `.github/workflows/release-security-report.yml`.
It runs Security and Malware scan on the tagged commit, merges the olevba,
mraptor, ClamAV and YARA-X results into one report, checks the report's
SHA-256 for each `.bas` against the release's own `.bas` files, and
attaches `security-report.md` and `security-report.json` to the release.
Releases from 3.8.3 on carry them. Started by hand, the workflow is a dry
run and attaches nothing.

### Verifying a download

The report lists every finding with its explanation and the SHA-256 of
each `.bas` build and of the workbook at the tag. Compare a download with
it:

```powershell
Get-FileHash .\ModernJsonInVBA_Excel.bas -Algorithm SHA256
```

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
