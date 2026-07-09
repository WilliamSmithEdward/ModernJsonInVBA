"""Combine the vba_source modules into single-file distributions in dist/.

Produces two files:

  dist/ModernJsonInVBA_Excel.bas     all eleven modules (parser + CSV/XML +
                                     Excel ListObject integration). Excel only.
  dist/ModernJsonInVBA_AllO365.bas   the nine host-agnostic modules (parsing,
                                     serialization, transforms, CSV, XML). No
                                     Excel references, so it imports into any
                                     O365 VBA host: Word, PowerPoint, Access,
                                     or Excel.

Merging VBA modules is not concatenation. This script:

  1. Keeps one Attribute VB_Name and one Option Explicit.
  2. Hoists every module-level declaration (Const / Type / Enum / module
     variable) above the first procedure, because VBA requires the whole
     declarations section to precede all procedures.
  3. Renames private module-level constants whose names collide across
     modules (for example ERR_SRC, which is "ModernJsonInVBA" in five modules
     and "XmlTextToJson" in Json_Xml). Values are preserved, so error sources
     do not change; only the private identifier is made unique per module.
  4. Fails loudly if two modules define a procedure with the same name.

The script also stamps the release version and date (read from the top
entry of CHANGELOG.md) into every module header in vba_source/ and into the
generated file headers, so the stamps cannot drift from the changelog.

Edit the modules in vba_source/ and re-run this script; do not edit the
generated files.

Usage:  python build_dist.py
"""
import os
import re
import sys

REPO = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(REPO, "vba_source")
DIST = os.path.join(REPO, "dist")
CHANGELOG = os.path.join(REPO, "CHANGELOG.md")
REPO_URL = "https://github.com/WilliamSmithEdward/ModernJsonInVBA"

ALL_O365_MODULES = [
    "Json_Common", "Json_Parser", "Json_Serializer", "Json_Model",
    "Json_Transforms", "Json_Tables", "Json_Coalesce", "Json_Csv", "Json_Xml",
    "Json_Ndjson",
]
EXCEL_MODULES = ALL_O365_MODULES + ["Json_Excel", "Json_Excel_Export"]

VB_NAME = "ModernJsonInVBA"

PROC_START = re.compile(
    r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*'
    r'(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w', re.IGNORECASE)
PROC_NAME = re.compile(
    r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*'
    r'(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)', re.IGNORECASE)
TYPE_START = re.compile(r'^\s*(?:Public\s+|Private\s+)?(?:Type|Enum)\s+\w', re.IGNORECASE)
TYPE_END = re.compile(r'^\s*End\s+(?:Type|Enum)\b', re.IGNORECASE)
PRIV_CONST = re.compile(r'^\s*Private\s+Const\s+(\w+)', re.IGNORECASE)
PRIV_VAR = re.compile(r'^\s*Private\s+(\w+)\s+As\s+', re.IGNORECASE)

EXCEL_TOKENS = re.compile(
    r'\b(?:Worksheet|ListObject|ListColumns?|ListRows?|Range|Application|'
    r'XlCalculation|xlSrcRange|xlYes|DataBodyRange|HeaderRowRange|Workbook|'
    r'Excel_[A-Za-z_]+)\b')


def read_lines(module):
    path = os.path.join(SRC, module + ".bas")
    with open(path, encoding="utf-8") as f:
        text = f.read()
    return text.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def read_release_info():
    """Version and date from the top release entry of CHANGELOG.md."""
    with open(CHANGELOG, encoding="utf-8") as f:
        for line in f:
            m = re.match(r"^## \[(\d+\.\d+\.\d+)\] - (\d{4}-\d{2}-\d{2})", line)
            if m:
                return m.group(1), m.group(2)
    raise SystemExit("no '## [x.y.z] - date' entry found in CHANGELOG.md")


def stamp_sources(version, released):
    """Refresh the Version/Released lines in every module header.

    Each module header carries 'Module:' and 'Project:' lines; the stamp
    lines sit directly below 'Project:'. Existing stamp lines are replaced,
    so the operation is idempotent and re-running after a release bump
    rewrites only the files whose stamps changed.
    """
    stamped = 0
    for module in EXCEL_MODULES:
        path = os.path.join(SRC, module + ".bas")
        with open(path, encoding="utf-8") as f:
            text = f.read().replace("\r\n", "\n").replace("\r", "\n")

        lines = []
        inserted = False
        for ln in text.split("\n"):
            if ln.startswith("' Version:") or ln.startswith("' Released:"):
                continue
            lines.append(ln)
            if not inserted and ln.startswith("' Project:"):
                lines.append("' Version:     " + version)
                lines.append("' Released:    " + released)
                inserted = True

        if not inserted:
            raise SystemExit(module + ": no \"' Project:\" header line to stamp under")

        new_text = "\n".join(lines)
        if new_text != text:
            with open(path, "w", encoding="utf-8", newline="") as f:
                f.write(new_text)
            stamped += 1

    print("stamped {} module header(s) at {} ({})".format(stamped, version, released))


def strip_header(lines):
    """Drop the Attribute VB_Name line and every Option Explicit line."""
    out = []
    for ln in lines:
        s = ln.strip()
        if s.lower().startswith("attribute vb_name"):
            continue
        if s.lower() == "option explicit":
            continue
        out.append(ln)
    return out


def split_decls_procs(lines):
    """Return (declaration lines, procedure lines).

    The split is the first module-level procedure. Comments, blank lines,
    Const/variable lines, and Type/Enum blocks before it are declarations.
    """
    i, n = 0, len(lines)
    in_block = False
    first_proc = n
    while i < n:
        s = lines[i].strip()
        if in_block:
            if TYPE_END.match(s):
                in_block = False
            i += 1
            continue
        if s == "" or s.startswith("'"):
            i += 1
            continue
        if TYPE_START.match(s):
            in_block = True
            i += 1
            continue
        if PROC_START.match(s):
            first_proc = i
            break
        i += 1  # some other declaration line
    return lines[:first_proc], lines[first_proc:]


def split_leading_comments(decls):
    """Separate the leading comment/blank banner from real declarations."""
    i = 0
    while i < len(decls) and (decls[i].strip() == "" or decls[i].strip().startswith("'")):
        i += 1
    return decls[:i], decls[i:]


def private_names(decls):
    names = []
    for ln in decls:
        s = ln.strip()
        m = PRIV_CONST.match(s)
        if m:
            names.append(m.group(1))
            continue
        m = PRIV_VAR.match(s)
        if m and m.group(1).lower() != "const":
            names.append(m.group(1))
    return names


def rename_token(lines, old, new):
    pat = re.compile(r'\b' + re.escape(old) + r'\b')
    return [pat.sub(new, ln) for ln in lines]


def trim_blanks(lines):
    while lines and lines[0].strip() == "":
        lines.pop(0)
    while lines and lines[-1].strip() == "":
        lines.pop()
    return lines


def build(modules, out_name, version, released):
    parsed = {}      # module -> dict(decls_head, decls_body, procs)
    priv_by_name = {}  # private name -> list of modules declaring it

    for m in modules:
        decls, procs = split_decls_procs(strip_header(read_lines(m)))
        head, body = split_leading_comments(decls)
        parsed[m] = {"head": head, "body": body, "procs": procs}
        for name in private_names(decls):
            priv_by_name.setdefault(name, []).append(m)

    # Rename any private module-level name that appears in more than one
    # module, so the merged single scope has no duplicate declarations.
    renames = {}  # (module, old) -> new
    for name, mods in priv_by_name.items():
        if len(mods) > 1:
            for m in mods:
                renames[(m, name)] = "{}_{}".format(m, name)

    for (m, old), new in renames.items():
        parsed[m]["body"] = rename_token(parsed[m]["body"], old, new)
        parsed[m]["procs"] = rename_token(parsed[m]["procs"], old, new)

    # Fail loudly on duplicate procedure names.
    seen_proc = {}
    for m in modules:
        for ln in parsed[m]["procs"]:
            pm = PROC_NAME.match(ln.strip())
            if pm:
                name = pm.group(1)
                if name.lower() in seen_proc:
                    raise SystemExit(
                        "duplicate procedure '{}' in {} and {}".format(
                            name, seen_proc[name.lower()], m))
                seen_proc[name.lower()] = m

    # Emit.
    out = []
    out.append('Attribute VB_Name = "{}"'.format(VB_NAME))
    out.append("Option Explicit")
    out.append("")
    out.append("' " + "=" * 76)
    out.append("' ModernJsonInVBA - single-file distribution")
    out.append("'")
    out.append("' Version:     " + version)
    out.append("' Released:    " + released)
    out.append("' Repo:        " + REPO_URL)
    out.append("'")
    out.append("' GENERATED by build_dist.py from vba_source/. Do not edit by hand;")
    out.append("' edit the modules in vba_source/ and re-run build_dist.py.")
    out.append("'")
    out.append("' Modules combined: " + ", ".join(modules))
    out.append("' " + "=" * 76)
    out.append("")

    out.append("' " + "=" * 76)
    out.append("' DECLARATIONS (hoisted above all procedures, as VBA requires)")
    out.append("' " + "=" * 76)
    for m in modules:
        body = trim_blanks(list(parsed[m]["body"]))
        if body:
            out.append("")
            out.append("' ---- {} ----".format(m))
            out.extend(body)
    out.append("")

    for m in modules:
        out.append("")
        out.append("' " + "=" * 76)
        out.append("' MODULE: {}".format(m))
        out.append("' " + "=" * 76)
        head = trim_blanks(list(parsed[m]["head"]))
        if head:
            out.extend(head)
            out.append("")
        out.extend(trim_blanks(list(parsed[m]["procs"])))

    text = "\r\n".join(out) + "\r\n"

    os.makedirs(DIST, exist_ok=True)
    path = os.path.join(DIST, out_name)
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(text)

    renamed = sorted({old for (_, old) in renames})
    print("wrote {}  ({} modules, {} lines)".format(
        os.path.relpath(path, REPO), len(modules), len(out)))
    if renamed:
        print("  collision-renamed private consts: " + ", ".join(renamed))
    return path


def main():
    version, released = read_release_info()
    stamp_sources(version, released)

    excel_path = build(EXCEL_MODULES, "ModernJsonInVBA_Excel.bas", version, released)
    o365_path = build(ALL_O365_MODULES, "ModernJsonInVBA_AllO365.bas", version, released)

    # Portability guard: the all-O365 file must reference no Excel object
    # model. String literals and comments are stripped first, so an Excel
    # word inside an error message (for example "you likely passed
    # Range.Value2") does not count.
    str_lit = re.compile(r'"(?:[^"]|"")*"')
    with open(o365_path, encoding="utf-8") as f:
        offenders = {}
        for i, ln in enumerate(f, 1):
            code = str_lit.sub('', ln).split("'", 1)[0]
            for tok in EXCEL_TOKENS.findall(code):
                offenders.setdefault(tok, []).append(i)
    if offenders:
        print("PORTABILITY CHECK FAILED: Excel references in all-O365 file:")
        for tok, lns in sorted(offenders.items()):
            print("  {} at lines {}".format(tok, lns[:5]))
        sys.exit(1)
    print("portability check: no Excel references in all-O365 file")


if __name__ == "__main__":
    main()
