# Security

## Reporting a vulnerability

Please report security problems privately through
[GitHub's private vulnerability reporting](https://github.com/WilliamSmithEdward/ModernJsonInVBA/security/advisories/new)
rather than in a public issue. Include the release you used, the host
application (Excel, Word, Access, PowerPoint), and a payload or macro that
reproduces the problem.

Fixes ship in a new release. Only the latest release is supported.

## What the library code does

The library is the twelve modules in `vba_source/` and the two single-file
builds in `dist/` that are generated from them. The code is plain VBA:

- No `Declare` statements, so no Windows or Mac API calls.
- No `CreateObject` or `GetObject`, so no COM objects, no scripting host,
  and no network access. The HTTP helpers in the README are examples for
  your own code; they are not part of the library.
- No `Shell`, no auto-run procedures (`AutoOpen`, `Workbook_Open`, and so
  on), and no reads of environment variables.
- One file operation: `Json_ReadTextFile` (also used by the CSV file
  reader) opens the path you pass it `For Binary Access Read`. Nothing in
  the library writes, deletes, or lists files.
- The Excel modules write only to the worksheet and table you pass in.

## Loading JSON you do not trust

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

## The example workbook

`ModernJsonInVBA.xlsm` contains the library plus test suites and examples.
Some of those macros download sample JSON from public demo APIs
(pokeapi.co, dummyjson.com, jsonplaceholder.typicode.com) with
`MSXML2.XMLHTTP`, and some write and delete temporary files in `%TEMP%`.
None of them run automatically. If you only need the library, import a
file from `dist/` instead of using the workbook.

## Automated scanning

Every push, pull request, release, and daily scheduled run scans every
VBA file in the repository (`.github/workflows/security.yml`):

- [olevba](https://github.com/decalage2/oletools/wiki/olevba) flags the
  keywords malicious macros depend on: file and process access, COM
  objects, auto-run entry points, encoded strings, and URLs.
- [mraptor](https://github.com/decalage2/oletools/wiki/mraptor) looks for
  the combination malicious macros need: an auto-run trigger plus a way to
  write files or execute code. It rates the Excel build and the workbook
  SUSPICIOUS on three false matches, all explained in the baseline: the
  public procedure `Excel_ResizeTableToRowCol` fits its pattern for event
  handlers like `UserForm_Resize`, it counts the read-only
  `Open ... For Binary Access Read` in `Json_ReadTextFile` as a write, and
  it reads "run" in comments as an execute. The per-file verdict is shown
  in the report; each individual match is checked against the baseline.

Plain code trips some of these on ordinary words in comments, such as "open"
in "open-addressing". Every accepted finding is therefore listed in
[`security/baseline.json`](security/baseline.json) with a note explaining
it. The library, the test suites, and the workbook each have their own list,
so a finding accepted for the tests is still a failure in the library. The
scan fails on any finding or URL host not on the list for its group.

The workflow shows ClamAV and YARA-X in their own CI job. It runs
[ClamAV](https://docs.clamav.net/manual/Usage/Scanning.html)
with official signatures refreshed by `freshclam` on each CI run and
[YARA-X](https://virustotal.github.io/yara-x/docs/api/python/) with the
pinned [YARA Forge core collection](https://github.com/YARAHQ/yara-forge/releases).
It scans the twelve source modules, both generated builds, test modules,
payload modules, and the shipping workbook. ClamAV also scans VBA extracted
from those files, and YARA-X runs focused local rules over that extracted
source. YARA Forge aggregates public rules from many authors; core favors
higher quality rules over the larger hunting sets. The report records scanner
versions, the collection URL and
archive SHA-256, every match, and scan errors. The download must match the
pinned archive SHA-256 before its rules are compiled. An unavailable scanner, failed
signature update, or rule compilation error fails CI.

The pinned YARA Forge release and SHA-256 are in
[`security/yara-forge.json`](security/yara-forge.json). A weekly workflow
(`.github/workflows/update-yara-forge.yml`) checks the latest stable release,
downloads the core archive, verifies its published checksum, and proposes
the new pin in a pull request. It dispatches the existing Security workflow
on the draft PR branch so that the new rules scan the repository before review.
The updater does not accept findings or merge the PR. Check the Security run
before marking the draft ready or merging it.

The local YARA-X rules mirror ReDim's focused checks for encoded PowerShell,
remote execution through Windows binaries, and Office Run key persistence.
They complement the public collection and have no standing exceptions.

Signature matches fail CI unless the exact scanner, signature, file path,
and file SHA-256 appear in [`security/malware-exceptions.json`](security/malware-exceptions.json)
with a reviewed reason. The file hash prevents an exception from covering
changed content. To review a new match, inspect the matched file and rule,
verify the source of the rule, then add only that tuple and a specific reason.
Do not bypass a scanner failure or allow an entire rule collection. Exceptions
no longer seen fail the scan until they are removed.

Releases from 3.8.3 onward carry `security-report.md` and
`security-report.json` as assets, attached by
`.github/workflows/release-security-report.yml` when the release is
published. The report lists every finding with its explanation and the
SHA-256 of each release file, so you can check that a downloaded `.bas` or
`.xlsm` matches the one that was scanned:

```powershell
Get-FileHash .\ModernJsonInVBA_Excel.bas -Algorithm SHA256
```

To run the same scan locally:

```bash
python -m pip install -r security/requirements.txt
python security/scan.py
```

To run the full CI scan locally, install ClamAV, update its signatures with
`freshclam`, then run (set `CLAMAV_DATABASE` if its database is in a
nondefault directory):

```bash
python security/malware_scan.py --out security-report/malware-results.json
python security/scan.py --malware-results security-report/malware-results.json
```

### Limits of the scan

All four scanners inspect static content. They do not run the code,
and a clean report does not prove the code is safe. A second workflow
(`.github/workflows/vba-analysis.yml`) runs
[pyVBAanalysis](https://github.com/WilliamSmithEdward/pyVBAanalysis) for
compile, type, and dead-code errors; it checks correctness, not intent.
Reading the source is the stronger check, and the library is short enough
to read.

olevba raises a "VBA Stomping" flag on the workbook. Stomping means the
compiled p-code Excel runs differs from the source you can read. olevba
compares the names and strings in the p-code with the source and stops at
the first one missing. On this workbook the misses are type suffixes that
the p-code dump adds (`esc$` for a name the source spells `esc`). That flag
is not allowlisted. Instead, `security/scan.py` repeats the comparison for
every name in the p-code and fails the scan on any miss other than such a
suffix. Identifiers match without regard to case, as VBA treats them.

That check found real residue in the workbook shipped with 3.8.2: three
names from deleted code (`Module1`, `PP_SortNames`, `tableRootLabel`) in
the p-code. They were leftovers from earlier edits, not hidden code, and
the workbook's VBA project has since been rebuilt from its module sources.
