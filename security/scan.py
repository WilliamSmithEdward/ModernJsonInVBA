"""Scan every VBA file in the repo with olevba and mraptor, against a baseline.

olevba (from oletools) flags VBA keywords that malware uses: file I/O,
process launch, network objects, auto-run entry points, obfuscated strings,
and URLs. mraptor (also from oletools) looks for the combination malicious
macros need: an auto-run trigger plus a way to write files or execute
code. Plain library code trips both tools on ordinary words in comments
("open-addressing", "call", "run"), so every accepted finding is listed in
security/baseline.json with a note saying why it is there.

mraptor reports one SUSPICIOUS/OK verdict per file. That verdict is recorded
but not gated; each individual mraptor match (its AutoExec, Write, and
Execute pattern hits) is a finding checked against the baseline like an
olevba finding, so any new trigger, write, or execute fails the scan.

Files are scanned in three groups with separate allowlists, so a finding
accepted for the test suites or the example workbook does not become
acceptable in the library:

  library    vba_source/*.bas, dist/*.bas
  tests      Tests/*.vba, json_payloads/*.bas
  workbook   ModernJsonInVBA.xlsm

The scan fails when a group shows a finding its allowlist does not contain,
a URL whose host is not allowed for that group, a tracked VBA file outside
every group, an accepted finding with no note, or an accepted finding the
group no longer shows. A stale entry must be removed from the baseline, so
the list stays the set of findings someone has actually reviewed.

olevba's "VBA Stomping" flag is never allowlisted. olevba compares the
names and strings in the compiled p-code with the source text and stops at
the first one the source lacks. pcodedmp renders some identifiers with a
type-declaration suffix the source never spells (initialCapacity$ for a
Long parameter), so an honest workbook trips it. When olevba raises the
flag, this script repeats the comparison over every name and accepts only
misses that differ from a source identifier by that one suffix character;
anything else fails the scan as possible stomping.

Writes security-report.md and security-report.json to --out.

Usage:
  python security/scan.py [--out DIR] [--label TEXT]
  python security/scan.py --update-baseline

--update-baseline replaces each group's accepted findings with what this
scan observed and adds "TODO" notes for new ones; the next scan fails until
every TODO is replaced with a real explanation. Review the diff before
committing it.
"""
import argparse
import datetime
import fnmatch
import glob
import hashlib
import json
import os
import re
import subprocess
import sys
from urllib.parse import urlsplit

from oletools import mraptor, olevba

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASELINE = os.path.join(REPO, "security", "baseline.json")

GROUPS = {
    "library": ["vba_source/*.bas", "dist/*.bas"],
    "tests": ["Tests/*.vba", "json_payloads/*.bas"],
    "workbook": ["ModernJsonInVBA.xlsm"],
}
# Tracked files with these extensions must belong to a group.
VBA_PATTERNS = ["*.bas", "*.cls", "*.frm", "*.vba", "*.xls", "*.xlsm", "*.xlsb",
                "*.xlam", "*.xla", "*.doc", "*.docm", "*.dotm", "*.ppt", "*.pptm"]
# Hashed in the report so a downloaded release file can be checked against it.
RELEASE_FILES = ["dist/*.bas", "ModernJsonInVBA.xlsm"]


def git(*args):
    return subprocess.run(["git", *args], cwd=REPO, capture_output=True,
                          text=True, check=True).stdout.strip()


def expand(patterns):
    out = []
    for p in patterns:
        hits = sorted(os.path.relpath(h, REPO).replace(os.sep, "/")
                      for h in glob.glob(os.path.join(REPO, p)))
        if not hits:
            raise SystemExit(f"scan: pattern {p!r} matched no files")
        out.extend(hits)
    return out


def untracked_by_groups(grouped):
    covered = {f for files in grouped.values() for f in files}
    tracked = git("ls-files").splitlines()
    return sorted(f for f in tracked
                  if any(fnmatch.fnmatch(f.lower(), p) for p in VBA_PATTERNS)
                  and f not in covered)


def finding_key(kind, keyword):
    # olevba reports the case it matched ("Open" in code, "open" in a
    # comment); case carries no security meaning, so keys are lowercased.
    return f"{kind}: {keyword}".lower()


STOMPING = ("Suspicious", "VBA Stomping")
# Type-declaration characters: String, Integer, Long, Single, Double,
# Currency, LongLong.
TYPE_SUFFIXES = "$%&!#@^"
STOMPING_NOTE = ("Checked by scan.py: every p-code name and string appears in the "
                 "source, apart from type-suffix renderings such as esc$.")


def pcode_names(parser):
    """Names and string literals in the p-code, extracted as olevba 0.60.2's
    detect_vba_stomping extracts them."""
    names = set()
    for line in parser.pcodedmp_output.splitlines():
        if not line.startswith("\t"):
            continue
        tokens = line.split(None, 1)
        mnemonic, args = tokens[0], (tokens[1].strip() if len(tokens) == 2 else "")
        if mnemonic in ("ArgsCall", "ArgsLd", "St", "Ld", "MemSt", "Label"):
            if args.startswith("(Call) "):
                args = args[7:]
            name = args.split(None, 1)[0]
            if not name.startswith("id_"):
                names.add(name)
        elif mnemonic == "LitStr":
            s = args.split(None, 1)[1]
            if len(s) >= 2:
                s = '"' + s[1:-1].replace('"', '""') + '"'
            names.add(s)
    return names


def stomping_residue(parser):
    """P-code names absent from the source, other than suffix renderings.

    String literals must match exactly. Identifiers match case-insensitively,
    as VBA does: the editor shows one spelling per name project-wide, so a
    parameter called "format" makes the source read format$ where the p-code
    says Format$."""
    code = parser.get_vba_code_all_modules()
    residue = []
    for name in sorted(pcode_names(parser)):
        if name in code:
            continue
        if not name.startswith('"'):
            base = name[:-1] if name[-1:] in TYPE_SUFFIXES else name
            if base and re.search(r"(?<![\w])" + re.escape(base) + r"(?![\w])", code, re.IGNORECASE):
                continue
        residue.append(name)
    return residue


MRAPTOR_PATTERNS = [
    ("mraptor AutoExec", mraptor.re_autoexec, "Matches mraptor's auto-run trigger pattern"),
    ("mraptor Write", mraptor.re_write, "Matches mraptor's file-write pattern"),
    ("mraptor Execute", mraptor.re_execute, "Matches mraptor's execute pattern"),
]


def mraptor_findings(code):
    """(verdict flags, [(kind, keyword, description)]) for one file's code.

    MacroRaptor.scan keeps only the first hit of each pattern; every hit is
    collected here so a new one cannot hide behind an accepted one."""
    raptor = mraptor.MacroRaptor(code)
    raptor.scan()
    verdict = ("SUSPICIOUS " if raptor.suspicious else "OK ") + raptor.get_flags()
    found = []
    for kind, pattern, description in MRAPTOR_PATTERNS:
        for m in pattern.finditer(code):
            found.append((kind, " ".join(m.group().split()), description))
    return verdict, found


def analyze(path):
    """(olevba file type, [(kind, keyword, description)], stomping residue,
    mraptor verdict) for one file. The residue is None unless olevba flagged
    stomping."""
    parser = olevba.VBA_Parser(os.path.join(REPO, path))
    try:
        if not parser.detect_vba_macros():
            return parser.type, [], None, "no macros"
        results = [tuple(r) for r in (parser.analyze_macros() or [])]
        residue = None
        if any((k, w) == STOMPING for k, w, _ in results):
            residue = stomping_residue(parser)
        verdict, raptor = mraptor_findings(parser.get_vba_code_all_modules())
        return parser.type, results + raptor, residue, verdict
    finally:
        parser.close()


def host_allowed(url, hosts):
    host = (urlsplit(url).hostname or "").lower()
    return any(host == h or host.endswith("." + h) for h in hosts)


def sha256(path):
    h = hashlib.sha256()
    with open(os.path.join(REPO, path), "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def scan(baseline):
    grouped = {g: expand(p) for g, p in GROUPS.items()}
    problems = []
    for f in untracked_by_groups(grouped):
        problems.append(f"{f}: tracked VBA file is not in any scan group")

    notes = baseline["notes"]
    groups = {}
    for group, files in grouped.items():
        policy = baseline["groups"][group]
        accepted = set(policy["accepted"])
        hosts = [h.lower() for h in policy["ioc_hosts"]]
        seen = {}  # key -> {"kind", "keyword", "description", "files"}
        residues = {}  # file -> stomping residue, for files olevba flagged
        verdicts = {}  # file -> mraptor verdict
        for f in files:
            ftype, results, residue, verdicts[f] = analyze(f)
            if ftype is None:
                problems.append(f"{f}: olevba could not identify the file type")
            if residue is not None:
                residues[f] = residue
            for kind, keyword, description in results:
                key = finding_key(kind, keyword)
                entry = seen.setdefault(key, {"kind": kind, "keyword": keyword,
                                              "description": description, "files": []})
                if f not in entry["files"]:
                    entry["files"].append(f)
        for key, entry in sorted(seen.items()):
            if entry["kind"] == "IOC" and entry["keyword"].lower().startswith(("http://", "https://")):
                ok = host_allowed(entry["keyword"], hosts)
                entry["status"] = "allowed host" if ok else "UNEXPECTED host"
            elif key == finding_key(*STOMPING):
                bad = {f: r for f, r in residues.items() if r}
                ok = not bad
                entry["status"] = "checked" if ok else "UNEXPECTED"
                entry["residue"] = bad
                for f, r in bad.items():
                    problems.append(f"{group}: possible VBA stomping in {f}: p-code names not in "
                                    f"the source: {', '.join(r[:10])}{' ...' if len(r) > 10 else ''}")
                entry["note"] = STOMPING_NOTE
                continue
            else:
                ok = key in accepted
                entry["status"] = "accepted" if ok else "UNEXPECTED"
            entry["note"] = notes.get(key, "") if entry["kind"] != "IOC" else ""
            if not ok:
                problems.append(f"{group}: unexpected {entry['kind']} {entry['keyword']!r} "
                                f"in {', '.join(entry['files'])}")
        stale = sorted(accepted - set(seen))
        for key in stale:
            problems.append(f"{group}: accepted finding {key!r} is no longer seen; "
                            "remove it from security/baseline.json")
        groups[group] = {"files": files, "findings": seen, "no_longer_seen": stale,
                         "mraptor": verdicts}

    for group, policy in baseline["groups"].items():
        for key in policy["accepted"]:
            if key == finding_key(*STOMPING):
                problems.append(f"baseline: {key!r} ({group}) cannot be accepted; "
                                "scan.py checks it on every run")
                continue
            note = notes.get(key, "")
            if not note or note.startswith("TODO"):
                problems.append(f"baseline: accepted finding {key!r} ({group}) has no note")
    return groups, problems


def update_baseline(baseline, groups):
    for group, data in groups.items():
        baseline["groups"][group]["accepted"] = sorted(
            k for k, e in data["findings"].items()
            if e["kind"] != "IOC" and k != finding_key(*STOMPING))
        for k in baseline["groups"][group]["accepted"]:
            baseline["notes"].setdefault(k, "TODO: explain why this is expected")
    used = {k for p in baseline["groups"].values() for k in p["accepted"]}
    baseline["notes"] = {k: v for k, v in sorted(baseline["notes"].items()) if k in used}
    with open(BASELINE, "w", encoding="utf-8", newline="\n") as f:
        json.dump(baseline, f, indent=2)
        f.write("\n")
    todo = [k for k, v in baseline["notes"].items() if v.startswith("TODO")]
    print(f"scan: baseline rewritten; {len(todo)} note(s) need an explanation")
    for k in todo:
        print("  TODO", k)


def md_cell(text):
    return str(text).replace("|", "\\|").replace("\n", " ")


def write_report(out_dir, label, groups, problems):
    os.makedirs(out_dir, exist_ok=True)
    commit = git("rev-parse", "HEAD")
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    result = "PASS" if not problems else "FAIL"
    hashes = [(f, sha256(f)) for f in expand(RELEASE_FILES)]

    lines = [
        f"# Security report: ModernJsonInVBA {label}",
        "",
        f"- Result: **{result}**",
        f"- Commit: `{commit}`",
        f"- Generated: {now}",
        f"- Scanners: olevba and mraptor from oletools {olevba.__version__}",
        "- Method and limits: see SECURITY.md in the repository",
        "",
        "## Release files (SHA-256)",
        "",
        "| File | SHA-256 |",
        "| --- | --- |",
    ]
    lines += [f"| `{f}` | `{h}` |" for f, h in hashes]
    if problems:
        lines += ["", "## Unexpected findings", ""]
        lines += [f"- {md_cell(p)}" for p in problems]
    for group, data in groups.items():
        lines += ["", f"## {group}", "",
                  "mraptor verdict per file (A = auto-run trigger, W = write, X = execute; "
                  "each match is listed in the table below):", "",
                  "| File | mraptor |", "| --- | --- |"]
        lines += [f"| `{f}` | {v} |" for f, v in data["mraptor"].items()]
        lines.append("")
        if not data["findings"]:
            lines.append("No findings.")
        else:
            lines += ["| Tool / type | Match | Status | Why it is expected | Files |",
                      "| --- | --- | --- | --- | --- |"]
            for e in data["findings"].values():
                files = e["files"] if len(e["files"]) <= 3 else e["files"][:3] + [f"+{len(e['files']) - 3} more"]
                lines.append(f"| {md_cell(e['kind'])} | `{md_cell(e['keyword'])}` | {e['status']} "
                             f"| {md_cell(e['note'])} | {md_cell(', '.join(files))} |")
        if data["no_longer_seen"]:
            lines += ["", "Accepted but no longer seen: "
                      + ", ".join(f"`{k}`" for k in data["no_longer_seen"])]
    lines.append("")

    with open(os.path.join(out_dir, "security-report.md"), "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines))
    report = {"label": label, "commit": commit, "generated": now, "result": result,
              "scanner": {"oletools": olevba.__version__},
              "release_files": dict(hashes), "problems": problems, "groups": groups}
    with open(os.path.join(out_dir, "security-report.json"), "w", encoding="utf-8", newline="\n") as f:
        json.dump(report, f, indent=2)
        f.write("\n")
    return result


def merge_malware_results(paths):
    """Combine the results of separate ClamAV and YARA-X runs into one."""
    merged = {"scanners": {}, "yara_forge_url": None, "yara_forge_sha256": None,
              "files": [], "modules": [], "findings": [], "problems": [],
              "exceptions_no_longer_seen": []}
    for path in paths:
        with open(path, encoding="utf-8") as f:
            part = json.load(f)
        merged["scanners"].update(part["scanners"])
        for key in ("yara_forge_url", "yara_forge_sha256"):
            merged[key] = merged[key] or part[key]
        for key in ("files", "modules"):
            merged[key] = sorted(set(merged[key]) | set(part[key]))
        for key in ("findings", "problems", "exceptions_no_longer_seen"):
            merged[key] += part[key]
    return merged


def add_malware_report(out_dir, malware_paths, result):
    """Merge signature scan results into both report formats."""
    malware = merge_malware_results(malware_paths)
    json_path = os.path.join(out_dir, "security-report.json")
    md_path = os.path.join(out_dir, "security-report.md")
    with open(json_path, encoding="utf-8") as f:
        report = json.load(f)
    report["malware"] = malware
    report["problems"].extend(malware["problems"])
    report["result"] = "FAIL" if report["problems"] else "PASS"
    with open(json_path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(report, f, indent=2)
        f.write("\n")
    with open(md_path, encoding="utf-8") as f:
        md = f.read()
    md = md.replace(f"- Result: **{result}**", f"- Result: **{report['result']}**", 1)
    lines = ["", "## ClamAV and YARA-X", "",
             f"- ClamAV: {malware['scanners'].get('clamav', 'unavailable')}",
             f"- YARA-X: {malware['scanners'].get('yara_x', 'unavailable')}",
             f"- YARA Forge core package SHA-256: `{malware['yara_forge_sha256']}`",
             "- Local rules: `security/vba_malware.yar` over extracted VBA",
             f"- Files scanned: {len(malware['files'])}",
             f"- VBA modules scanned: {len(malware['modules'])}", ""]
    if malware["findings"]:
        lines += ["| Scanner | Signature | File | SHA-256 | Status | Reason |",
                  "| --- | --- | --- | --- | --- | --- |"]
        for e in malware["findings"]:
            lines.append("| " + " | ".join(md_cell(e[k]) for k in
                         ("scanner", "signature", "file", "sha256", "status", "reason")) + " |")
    else:
        lines.append("No signature matches.")
    if malware["problems"]:
        lines += ["", "### Scan failures", ""]
        lines += [f"- {md_cell(p)}" for p in malware["problems"]]
    with open(md_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(md.rstrip() + "\n" + "\n".join(lines) + "\n")
    return report["result"]


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--out", default=os.path.join(REPO, "security-report"))
    ap.add_argument("--label", help="release or ref name for the report title")
    ap.add_argument("--update-baseline", action="store_true")
    ap.add_argument("--malware-results", action="append",
                    help="JSON from security/malware_scan.py; repeat for each scanner run")
    args = ap.parse_args()

    with open(BASELINE, encoding="utf-8") as f:
        baseline = json.load(f)
    groups, problems = scan(baseline)
    if args.update_baseline:
        update_baseline(baseline, groups)
        return 0

    label = args.label or git("describe", "--tags", "--always", "--dirty")
    result = write_report(args.out, label, groups, problems)
    if args.malware_results:
        result = add_malware_report(args.out, args.malware_results, result)
    for p in problems:
        print("scan:", p)
    print(f"scan: {result}; report in {args.out}")
    return 0 if result == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
