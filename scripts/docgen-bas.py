import re
from pathlib import Path

BAS_DIR = Path("src/bas")
OUT_FILE = Path("docs") / "bas.md"


def find_procedures(text):
    """Return list of {name, kind, body}."""
    lines = text.splitlines()
    procs = []

    decl_re = re.compile(
        r'^\s*(Public|Private|Friend|Global|Static)?\s*'
        r'(Sub|Function)\s+([A-Za-z0-9_]+)\s*\(',
        re.IGNORECASE,
    )

    i = 0
    n = len(lines)
    while i < n:
        m = decl_re.match(lines[i])
        if not m:
            i += 1
            continue

        kind = m.group(2).capitalize()   # Sub / Function
        name = m.group(3)
        start = i

        end_re = re.compile(r'^\s*End\s+%s\b' % kind, re.IGNORECASE)
        j = i + 1
        while j < n and not end_re.match(lines[j]):
            j += 1
        if j < n:
            j_end = j
        else:
            j_end = n - 1

        body = "\n".join(lines[start : j_end + 1])
        procs.append({"name": name, "kind": kind, "body": body})
        i = j_end + 1

    return procs


def main():
    modules = []

    for bas_path in sorted(BAS_DIR.glob("*.bas")):
        text = bas_path.read_text(encoding="utf-8", errors="ignore")
        procs = find_procedures(text)
        if procs:
            modules.append(
                {"file": bas_path.name, "module": bas_path.stem, "procs": procs}
            )

    lines = []
    lines.append("# VBA modules")
    lines.append("")
    lines.append(
        "_This file is generated automatically from `.bas` files in `src/bas`._"
    )
    lines.append("")

    for mod in modules:
        lines.append(f"## Module `{mod['module']}`")
        lines.append("")
        for proc in mod["procs"]:
            title = f"{proc['kind']} {proc['name']}"
            lines.append(f"### `{proc['name']}`")
            lines.append("")
            lines.append(f"```vbnet")
            lines.append(proc["body"])
            lines.append("```")
            lines.append("")

    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUT_FILE.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    main()
