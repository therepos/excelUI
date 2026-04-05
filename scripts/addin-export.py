"""
Purpose: Export VBA modules, ribbon XML, and .xlam from your Excel add-in.
Usage: Double-click to run.
"""

import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


# ─── Auto-install oletools if missing ────────────────────────────────────────

try:
    import oletools.olevba  # noqa: F401
except ImportError:
    print("oletools not found. Installing...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "oletools", "--quiet"])
    print("oletools installed.\n")


# ═════════════════════════════════════════════════════════════════════════════
# CONFIGURATION — add or remove scan paths as needed
# ═════════════════════════════════════════════════════════════════════════════

APPDATA = os.environ.get("APPDATA", "")

SCAN_PATHS = [
    (Path(APPDATA) / "Microsoft" / "AddIns",            [".xlam", ".ppam"]),
    (Path(APPDATA) / "Microsoft" / "Word" / "STARTUP",  [".dotm"]),
    (Path(APPDATA) / "Microsoft" / "PowerPoint" / "AddIns", [".ppam"]),
]

# ═════════════════════════════════════════════════════════════════════════════


APP_LABELS = {
    ".xlam": "Excel",
    ".dotm": "Word",
    ".ppam": "PowerPoint",
}


def script_dir() -> Path:
    return Path(__file__).resolve().parent


def scan_addins() -> list[dict]:
    """Scan all configured paths for Office add-in files."""
    found = []
    seen = set()

    for folder, extensions in SCAN_PATHS:
        if not folder.exists():
            continue
        for f in folder.iterdir():
            if f.suffix.lower() in extensions and f.resolve() not in seen:
                seen.add(f.resolve())
                found.append({
                    "path": f,
                    "name": f.stem,
                    "ext": f.suffix.lower(),
                    "app": APP_LABELS.get(f.suffix.lower(), "Office"),
                })

    # Sort by app type then name
    found.sort(key=lambda x: (x["app"], x["name"].lower()))
    return found


def extract_xml(addin_path: Path, output_dir: Path) -> bool:
    print("\n[1/3] Extracting ribbon XML...")

    candidates = [
        "customUI/customUI14.xml",
        "customUI/customUI.xml",
        "customUI14/customUI14.xml",
    ]

    try:
        with zipfile.ZipFile(addin_path, "r") as zf:
            for candidate in candidates:
                if candidate in zf.namelist():
                    output_dir.mkdir(parents=True, exist_ok=True)
                    dest = output_dir / (addin_path.stem + ".xml")
                    with zf.open(candidate) as src, open(dest, "wb") as dst:
                        dst.write(src.read())
                    print(f"  -> {dest}")
                    return True
    except zipfile.BadZipFile:
        print("  -> WARNING: File is not a valid ZIP/Office archive")
        return False

    print("  -> No customUI XML found (skipped)")
    return False


def extract_bas(addin_path: Path, output_dir: Path) -> bool:
    print("\n[2/3] Extracting VBA modules (.bas)...")

    try:
        with zipfile.ZipFile(addin_path, "r") as zf:
            if "xl/vbaProject.bin" not in zf.namelist() and \
               "word/vbaProject.bin" not in zf.namelist() and \
               "ppt/vbaProject.bin" not in zf.namelist():
                print("  -> No vbaProject.bin found (skipped)")
                return False
    except zipfile.BadZipFile:
        print("  -> WARNING: File is not a valid ZIP/Office archive")
        return False

    try:
        result = subprocess.run(
            [sys.executable, "-m", "oletools.olevba", "--decode", str(addin_path)],
            capture_output=True,
            text=True,
            timeout=30,
        )
        vba_output = result.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired) as e:
        print(f"  -> ERROR: {e}")
        return False

    if not vba_output.strip():
        print("  -> WARNING: olevba returned no output")
        return False

    sections = re.split(r"\n-{10,}\n", vba_output)

    skip = ("ThisWorkbook", "ThisDocument", "Sheet")
    output_dir.mkdir(parents=True, exist_ok=True)
    exported = 0

    for section in sections:
        match = re.match(r"VBA MACRO (\S+)", section.strip())
        if not match:
            continue

        name = match.group(1)

        if any(name.startswith(s) for s in skip):
            continue

        if not name.endswith(".bas"):
            continue

        code_match = re.search(r"- - - - .*?\n(.*)", section, re.DOTALL)
        if not code_match:
            continue

        code = code_match.group(1).rstrip()

        if not code or code.strip() == "(empty macro)":
            continue

        table_start = re.search(r"\n\+[-+]+\+\n", code)
        if table_start:
            code = code[:table_start.start()].rstrip()

        dest = output_dir / name
        dest.write_text(code + "\n", encoding="utf-8")
        print(f"  -> {dest}")
        exported += 1

    if exported == 0:
        print("  -> No standard modules found")
        return False

    print(f"  -> {exported} module(s) exported")
    return True


def copy_addin(addin_path: Path, output_dir: Path) -> bool:
    print("\n[3/3] Copying add-in file...")

    output_dir.mkdir(parents=True, exist_ok=True)
    dest = output_dir / addin_path.name

    shutil.copy2(addin_path, dest)
    print(f"  -> {dest}")
    return True


def main():
    print()
    print("=" * 40)
    print("  Office Add-in Export")
    print("=" * 40)

    addins = scan_addins()

    if not addins:
        print("\nNo add-ins found in the scanned folders:")
        for folder, exts in SCAN_PATHS:
            print(f"  {folder}  ({', '.join(exts)})")
        print()
        input("Press Enter to exit...")
        return

    print("\nFound add-ins:\n")
    for idx, a in enumerate(addins, 1):
        print(f"  [{idx}] {a['name']}{a['ext']}  ({a['app']})")

    print()
    choice = input("Enter number to export (or q to quit): ").strip()

    if choice.lower() == "q" or choice == "":
        print("Cancelled.")
        input("\nPress Enter to exit...")
        return

    try:
        idx = int(choice) - 1
        if idx < 0 or idx >= len(addins):
            raise ValueError
    except ValueError:
        print("Invalid selection.")
        input("\nPress Enter to exit...")
        return

    selected = addins[idx]
    addin_path = selected["path"]
    root = script_dir()
    output_dir = root / selected["name"]

    print()
    print(f"  Exporting: {selected['name']}{selected['ext']}  ({selected['app']})")
    print(f"  From:      {addin_path}")
    print(f"  To:        {output_dir}")

    extract_xml(addin_path, output_dir)
    extract_bas(addin_path, output_dir)
    copy_addin(addin_path, output_dir)

    print()
    print("=" * 40)
    print("  Done!")
    print("=" * 40)
    print()
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()