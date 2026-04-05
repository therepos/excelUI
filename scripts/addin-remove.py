"""
Purpose: Remove Office add-ins (Excel, Word, PowerPoint).
Usage: Double-click to run.
"""

import os
from pathlib import Path


# ═════════════════════════════════════════════════════════════════════════════
# SCAN PATHS — add or remove as needed
# ═════════════════════════════════════════════════════════════════════════════

APPDATA = os.environ.get("APPDATA", "")

SCAN_PATHS = [
    (Path(APPDATA) / "Microsoft" / "AddIns",            [".xlam", ".ppam"]),
    (Path(APPDATA) / "Microsoft" / "Word" / "STARTUP",  [".dotm"]),
    (Path(APPDATA) / "Microsoft" / "PowerPoint" / "AddIns", [".ppam"]),
]

APP_LABELS = {
    ".xlam": "Excel",
    ".dotm": "Word",
    ".ppam": "PowerPoint",
}

# ═════════════════════════════════════════════════════════════════════════════


def scan_addins() -> list[dict]:
    found = []
    seen = set()

    for folder, extensions in SCAN_PATHS:
        if not folder.exists():
            continue
        for f in folder.iterdir():
            if f.is_file() and f.suffix.lower() in extensions and f.resolve() not in seen:
                seen.add(f.resolve())
                found.append({
                    "name": f.name,
                    "path": f,
                    "app": APP_LABELS.get(f.suffix.lower(), "Office"),
                })

    found.sort(key=lambda x: (x["app"], x["name"].lower()))
    return found


def main():
    print()
    print("=" * 40)
    print("  Office Add-in Remover")
    print("=" * 40)

    addins = scan_addins()

    if not addins:
        print("\nNo add-ins found.")
        input("\nPress Enter to exit...")
        return

    print("\nFound add-ins:\n")
    for i, a in enumerate(addins, 1):
        print(f"  [{i}] {a['name']}  ({a['app']})")

    print()
    choice = input("Enter number to remove (or q to quit): ").strip()

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
    print(f"\n  File: {selected['path']}")
    confirm = input(f"\n  Delete {selected['name']}? (y/n): ").strip().lower()

    if confirm != "y":
        print("Cancelled.")
        input("\nPress Enter to exit...")
        return

    try:
        selected["path"].unlink()
        print(f"\n  Removed: {selected['name']}")
    except PermissionError:
        print(f"\n  ERROR: Cannot delete — is {selected['app']} still open?")
    except Exception as e:
        print(f"\n  ERROR: {e}")

    print()
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()
