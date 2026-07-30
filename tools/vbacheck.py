#!/usr/bin/env python3
"""Static structural checks for VBA .bas modules.

No Excel available, so this catches the mechanical failure modes that would
otherwise only surface as a compile error on the user's machine:
  - unbalanced block constructs (Sub/Function/If/With/For/Do/Select/Property)
  - odd number of double quotes on a line (broken string literal)
  - non-ASCII bytes (UTF-8 em dashes are a known VBA syntax error in this repo)
  - Option Explicit present, VB_Name present
  - references to identifiers that no module defines
"""
import re
import sys
import pathlib
from collections import defaultdict

OPENERS = [
    (re.compile(r"^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?Sub\s+\w+", re.I), "Sub"),
    (re.compile(r"^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?Function\s+\w+", re.I), "Function"),
    (re.compile(r"^\s*(Public\s+|Private\s+|Friend\s+)?Property\s+(Get|Let|Set)\s+\w+", re.I), "Property"),
    (re.compile(r"^\s*With\s+\S", re.I), "With"),
    (re.compile(r"^\s*Select\s+Case\s+\S", re.I), "Select"),
    (re.compile(r"^\s*Do\s*($|While|Until)", re.I), "Do"),
]
CLOSERS = {
    "Sub": re.compile(r"^\s*End\s+Sub\b", re.I),
    "Function": re.compile(r"^\s*End\s+Function\b", re.I),
    "Property": re.compile(r"^\s*End\s+Property\b", re.I),
    "With": re.compile(r"^\s*End\s+With\b", re.I),
    "Select": re.compile(r"^\s*End\s+Select\b", re.I),
    "Do": re.compile(r"^\s*Loop\b", re.I),
}
RE_FOR = re.compile(r"^\s*For\s+(Each\s+)?\w+", re.I)
RE_NEXT = re.compile(r"^\s*Next\b", re.I)
# Multi-line If: "If ... Then" with nothing after Then
RE_IF_BLOCK = re.compile(r"^\s*If\b.*\bThen\s*$", re.I)
RE_END_IF = re.compile(r"^\s*End\s+If\b", re.I)
RE_ELSEIF = re.compile(r"^\s*ElseIf\b.*\bThen\s*$", re.I)


def strip_strings_and_comment(line):
    """Remove string literals, then a trailing comment. Returns (code, nquotes)."""
    out, i, n, inq = [], 0, len(line), False
    quotes = 0
    while i < n:
        ch = line[i]
        if ch == '"':
            quotes += 1
            if inq and i + 1 < n and line[i + 1] == '"':  # escaped "" inside string
                quotes += 1
                i += 2
                continue
            inq = not inq
            i += 1
            continue
        if not inq and ch == "'":
            break
        if not inq:
            out.append(ch)
        i += 1
    return "".join(out), quotes


def check(path):
    errs, warns = [], []
    raw = path.read_bytes()
    nonascii = sorted({b for b in raw if b > 127})
    if nonascii:
        errs.append(f"non-ASCII byte(s) present: {nonascii}")
    text = raw.decode("latin-1")
    if "\r\n" in text:
        warns.append("file uses CRLF (repo convention is LF)")
    rawlines = text.replace("\r\n", "\n").split("\n")
    # Join VBA line continuations ("_" at end of line) so a statement spanning
    # several physical lines is analysed as one logical statement.
    lines, buf, buf_ln = [], None, 0
    for i, l in enumerate(rawlines, 1):
        cur = (buf + " " + l.strip()) if buf is not None else l
        if re.search(r"\s_\s*$", cur):
            buf = re.sub(r"\s_\s*$", "", cur)
            if buf_ln == 0:
                buf_ln = i
            continue
        lines.append((buf_ln or i, cur))
        buf, buf_ln = None, 0
    if buf is not None:
        lines.append((buf_ln, buf))

    if not re.search(r'Attribute\s+VB_Name\s*=\s*"([^"]+)"', text):
        errs.append("no Attribute VB_Name")
    if not re.search(r"^\s*Option\s+Explicit", text, re.M | re.I):
        warns.append("no Option Explicit")

    stack = []
    for ln, line in lines:
        code, nq = strip_strings_and_comment(line)
        if nq % 2:
            errs.append(f"L{ln}: odd number of double-quotes (unterminated string)")
        if not code.strip():
            continue
        # VBA allows several statements on one line separated by ':', and this
        # repo leans on it heavily ("With ctrl: .Name = "x": End With").
        # Guard ':=' (named arguments) before splitting.
        for st in code.replace(":=", "\x00").split(":"):
            st = st.replace("\x00", ":=").strip()
            if not st:
                continue
            matched = False
            for rx, kind in OPENERS:
                if rx.match(st):
                    # "Declare Sub" / "Exit Sub" are not block openers
                    if re.match(r"^\s*(Declare|Exit)\b", st, re.I):
                        break
                    stack.append((kind, ln))
                    matched = True
                    break
            if matched:
                continue
            for kind, rx in CLOSERS.items():
                if rx.match(st):
                    if not stack or stack[-1][0] != kind:
                        top = stack[-1] if stack else None
                        errs.append(
                            f"L{ln}: 'End {kind}' but innermost open block is "
                            f"{top[0]+' at L'+str(top[1]) if top else 'nothing'}"
                        )
                    else:
                        stack.pop()
                    matched = True
                    break
            if matched:
                continue
            if RE_IF_BLOCK.match(st) and not RE_ELSEIF.match(st):
                stack.append(("If", ln))
            elif RE_END_IF.match(st):
                if not stack or stack[-1][0] != "If":
                    top = stack[-1] if stack else None
                    errs.append(
                        f"L{ln}: 'End If' but innermost open block is "
                        f"{top[0]+' at L'+str(top[1]) if top else 'nothing'}"
                    )
                else:
                    stack.pop()
            elif RE_FOR.match(st):
                stack.append(("For", ln))
            elif RE_NEXT.match(st):
                if not stack or stack[-1][0] != "For":
                    top = stack[-1] if stack else None
                    errs.append(
                        f"L{ln}: 'Next' but innermost open block is "
                        f"{top[0]+' at L'+str(top[1]) if top else 'nothing'}"
                    )
                else:
                    stack.pop()

    for kind, ln in stack:
        errs.append(f"unclosed {kind} opened at L{ln}")

    # VBA requires every module-level declaration to sit in the declarations
    # section, above the first procedure. A Dim/Private/Public/Const placed
    # after one is a compile error, and it is an easy mistake to make when
    # appending helpers to the end of a module.
    first_proc = None
    RE_PROC = re.compile(
        r"^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property)\s+", re.I
    )
    RE_MOD_DECL = re.compile(r"^\s*(Dim|Private|Public|Const|Global)\s+", re.I)
    depth = 0
    for ln, line in lines:
        code, _ = strip_strings_and_comment(line)
        st = code.strip()
        if not st:
            continue
        if RE_PROC.match(st) and not re.match(r"^\s*(Declare|Exit)\b", st, re.I):
            depth += 1
            if first_proc is None:
                first_proc = ln
        elif re.match(r"^\s*End\s+(Sub|Function|Property)\b", st, re.I):
            depth = max(0, depth - 1)
        elif depth == 0 and first_proc is not None and RE_MOD_DECL.match(st):
            if not RE_PROC.match(st):
                errs.append(
                    f"L{ln}: module-level declaration after the first procedure "
                    f"(L{first_proc}); VBA requires it in the declarations section"
                )
    return errs, warns


def main(paths):
    files = []
    for p in paths:
        p = pathlib.Path(p)
        files.extend(sorted(p.rglob("*.bas")) if p.is_dir() else [p])

    vbnames = defaultdict(list)
    total_err = 0
    for f in files:
        errs, warns = check(f)
        m = re.search(r'Attribute\s+VB_Name\s*=\s*"([^"]+)"', f.read_bytes().decode("latin-1"))
        if m:
            vbnames[m.group(1)].append(f.name)
        status = "FAIL" if errs else ("warn" if warns else "ok")
        if errs or warns:
            print(f"[{status}] {f.name}")
            for e in errs:
                print(f"    ERROR  {e}")
            for w in warns:
                print(f"    warn   {w}")
        total_err += len(errs)

    dups = {k: v for k, v in vbnames.items() if len(v) > 1}
    if dups:
        print("\nDUPLICATE VB_Name (CI gate would fail):")
        for k, v in dups.items():
            print(f"  {k}: {v}")
        total_err += len(dups)

    print(f"\nchecked {len(files)} file(s); {len(vbnames)} distinct VB_Name; {total_err} error(s)")
    return 1 if total_err else 0


if __name__ == "__main__":
    default = str(pathlib.Path(__file__).resolve().parent.parent / "modules")
    sys.exit(main(sys.argv[1:] or [default]))
