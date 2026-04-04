"""
extract-vba-keys.py — Extract CMG/DPB/GC values from a password-protected VBA project.

Steps:
  1. Open Excel, create a new .xlsm file
  2. Alt+F11 → Tools → VBA Project Properties → Protection tab
  3. Check "Lock project for viewing", set password: Xk9#mTvL4@nRqW2f
  4. Save as .xlsm and close Excel
  5. Run: python extract-vba-keys.py <your_file.xlsm>
  6. Copy the output into lock-vba.py
"""

import re
import sys
import zipfile


def main():
    if len(sys.argv) < 2:
        print("Usage: python extract-vba-keys.py <file.xlsm>")
        sys.exit(1)

    path = sys.argv[1]

    with zipfile.ZipFile(path, "r") as zf:
        for name in ["xl/vbaProject.bin", "word/vbaProject.bin", "ppt/vbaProject.bin"]:
            if name in zf.namelist():
                data = zf.read(name)
                break
        else:
            print("ERROR: No vbaProject.bin found")
            sys.exit(1)

    text = data.decode("latin-1")

    keys = {}
    for key in ["ID", "CMG", "DPB", "GC"]:
        match = re.search(rf'{key}="([^"]*)"', text)
        if match:
            keys[key] = match.group(1)

    if len(keys) < 4:
        print("ERROR: Could not find all keys. Is the VBA project password-protected?")
        print(f"  Found: {', '.join(keys.keys())}")
        sys.exit(1)

    print()
    print("Found! Paste these into lock-vba.py:")
    print()
    print(f"LOCK_ID   = '{keys['ID']}'")
    print(f"LOCK_CMG  = '{keys['CMG']}'")
    print(f"LOCK_DPB  = '{keys['DPB']}'")
    print(f"LOCK_GC   = '{keys['GC']}'")
    print()


if __name__ == "__main__":
    main()