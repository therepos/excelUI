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
# CONFIGURATION — edit this path to point to your .xlam
# Leave empty to use the default: %APPDATA%\Microsoft\AddIns\excelEY.xlam
# ═════════════════════════════════════════════════════════════════════════════

XLAM_PATH = ""

# ═════════════════════════════════════════════════════════════════════════════


def script_dir() -> Path:
    return Path(__file__).resolve().parent


def get_xlam_path() -> Path:
    if XLAM_PATH:
        return Path(XLAM_PATH)
    appdata = os.environ.get("APPDATA", "")
    return Path(appdata) / "Microsoft" / "AddIns" / "excelEY.xlam"


def extract_xml(xlam_path: Path, output_dir: Path) -> bool:
    print("\n[1/3] Extracting ribbon XML...")

    candidates = [
        "customUI/customUI14.xml",
        "customUI/customUI.xml",
        "customUI14/customUI14.xml",
    ]

    with zipfile.ZipFile(xlam_path, "r") as zf:
        for candidate in candidates:
            if candidate in zf.namelist():
                output_dir.mkdir(parents=True, exist_ok=True)
                dest = output_dir / (xlam_path.stem + ".xml")
                with zf.open(candidate) as src, open(dest, "wb") as dst:
                    dst.write(src.read())
                print(f"  -> {dest}")
                return True

    print("  -> WARNING: No customUI XML found in .xlam")
    return False


def extract_bas(xlam_path: Path, output_dir: Path) -> bool:
    print("\n[2/3] Extracting VBA modules (.bas)...")

    with zipfile.ZipFile(xlam_path, "r") as zf:
        if "xl/vbaProject.bin" not in zf.namelist():
            print("  -> WARNING: No vbaProject.bin found")
            return False

    try:
        result = subprocess.run(
            [sys.executable, "-m", "oletools.olevba", "--decode", str(xlam_path)],
            capture_output=True,
            text=True,
            timeout=30,
        )
        vba_output = result.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired) as e:
        print(f"  -> ERROR: {e}")
        print("     Make sure oletools is installed: pip install oletools")
        return False

    if not vba_output.strip():
        print("  -> WARNING: olevba returned no output")
        return False

    # olevba output format:
    # -------------------------------------------------------------------------------
    # VBA MACRO EYRetain.bas
    # in file: xl/vbaProject.bin - OLE stream: 'VBA/EYRetain'
    # - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    # <code here>
    # -------------------------------------------------------------------------------
    # VBA MACRO next...

    # Split on the solid separator lines (----...----)
    sections = re.split(r"\n-{10,}\n", vba_output)

    skip = ("ThisWorkbook", "Sheet")
    output_dir.mkdir(parents=True, exist_ok=True)
    exported = 0

    for section in sections:
        # Look for module header
        match = re.match(r"VBA MACRO (\S+)", section.strip())
        if not match:
            continue

        name = match.group(1)

        # Skip ThisWorkbook, Sheet modules
        if any(name.startswith(s) for s in skip):
            continue

        # Only export .bas modules (skip .cls etc)
        if not name.endswith(".bas"):
            continue

        # Code starts after the "- - - - -" dashed line
        code_match = re.search(r"- - - - .*?\n(.*)", section, re.DOTALL)
        if not code_match:
            continue

        code = code_match.group(1).rstrip()

        # Skip empty macros
        if not code or code.strip() == "(empty macro)":
            continue

        # Remove the suspicious keywords table that olevba appends at the end
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


def copy_xlam(xlam_path: Path, output_dir: Path) -> bool:
    print("\n[3/3] Copying .xlam...")

    output_dir.mkdir(parents=True, exist_ok=True)
    dest = output_dir / xlam_path.name

    shutil.copy2(xlam_path, dest)
    print(f"  -> {dest}")
    return True


def main():
    xlam_path = get_xlam_path()
    root = script_dir()

    if not xlam_path.exists():
        print(f"ERROR: Cannot find .xlam at: {xlam_path}")
        print()
        print("Edit XLAM_PATH at the top of this script, e.g.:")
        print(r'  XLAM_PATH = r"C:\apps\myAddin.xlam"')
        sys.exit(1)

    print()
    print("=" * 40)
    print("  excelEY Export")
    print("=" * 40)
    print(f"  From:  {xlam_path}")
    print(f"  To:    {root}")

    extract_xml(xlam_path, root / "xml")
    extract_bas(xlam_path, root / "bas")
    copy_xlam(xlam_path, root / "xlam")

    print()
    print("=" * 40)
    print("  Done!")
    print("=" * 40)
    print()


if __name__ == "__main__":
    main()
    input("Press Enter to exit...")