"""
Purpose: Cleanly remove Office add-ins (Excel, Word, PowerPoint).
         Deletes files AND removes registry/config references so Office
         won't show "Sorry, we couldn't find ..." errors on startup.
Usage:   Double-click to run  (or:  python addin-remove.py)
Note:    Run with the target Office apps CLOSED for best results.
"""

import os
import sys
import winreg
from pathlib import Path


# ═════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═════════════════════════════════════════════════════════════════════════════

APPDATA = os.environ.get("APPDATA", "")

# Folders to scan for add-in files
SCAN_PATHS = [
    (Path(APPDATA) / "Microsoft" / "AddIns",                [".xlam", ".xla", ".ppam", ".ppa"]),
    (Path(APPDATA) / "Microsoft" / "Word" / "STARTUP",      [".dotm", ".dot"]),
    (Path(APPDATA) / "Microsoft" / "PowerPoint" / "AddIns", [".ppam", ".ppa"]),
    (Path(APPDATA) / "Microsoft" / "Excel" / "XLSTART",     [".xlam", ".xla", ".xls", ".xlsx"]),
]

APP_LABELS = {
    ".xlam": "Excel",  ".xla": "Excel",
    ".dotm": "Word",   ".dot": "Word",
    ".ppam": "PowerPoint", ".ppa": "PowerPoint",
    ".xls": "Excel",   ".xlsx": "Excel",
}

# Registry locations where Office registers add-ins
# Each entry: (registry key path, app filter or None for all)
EXCEL_ADDIN_REG_KEYS = [
    (r"Software\Microsoft\Office\Excel\Addins", None),
    (r"Software\Microsoft\Office\16.0\Excel\Options", None),
    (r"Software\Microsoft\Office\15.0\Excel\Options", None),
    (r"Software\Microsoft\Office\14.0\Excel\Options", None),
]

WORD_ADDIN_REG_KEYS = [
    (r"Software\Microsoft\Office\Word\Addins", None),
    (r"Software\Microsoft\Office\16.0\Word\Options", None),
    (r"Software\Microsoft\Office\15.0\Word\Options", None),
]

PPT_ADDIN_REG_KEYS = [
    (r"Software\Microsoft\Office\PowerPoint\Addins", None),
    (r"Software\Microsoft\Office\16.0\PowerPoint\Options", None),
    (r"Software\Microsoft\Office\15.0\PowerPoint\Options", None),
]

HIVE = winreg.HKEY_CURRENT_USER


# ═════════════════════════════════════════════════════════════════════════════
# SCANNING
# ═════════════════════════════════════════════════════════════════════════════

def scan_addin_files() -> list[dict]:
    """Find add-in files on disk."""
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
                    "source": "file",
                })

    found.sort(key=lambda x: (x["app"], x["name"].lower()))
    return found


def scan_registry_addins() -> list[dict]:
    """Find add-in references in the Windows Registry."""
    found = []

    all_reg_keys = EXCEL_ADDIN_REG_KEYS + WORD_ADDIN_REG_KEYS + PPT_ADDIN_REG_KEYS

    for key_path, _ in all_reg_keys:
        try:
            key = winreg.OpenKey(HIVE, key_path, 0, winreg.KEY_READ)
        except FileNotFoundError:
            continue
        except OSError:
            continue

        # Check for COM add-in subkeys (e.g. Office\Excel\Addins\SomeAddin)
        try:
            i = 0
            while True:
                subkey_name = winreg.EnumKey(key, i)
                try:
                    subkey = winreg.OpenKey(key, subkey_name, 0, winreg.KEY_READ)
                    manifest = ""
                    try:
                        manifest, _ = winreg.QueryValueEx(subkey, "Manifest")
                    except FileNotFoundError:
                        try:
                            manifest, _ = winreg.QueryValueEx(subkey, "Path")
                        except FileNotFoundError:
                            pass
                    winreg.CloseKey(subkey)

                    app = "Office"
                    if "Excel" in key_path:
                        app = "Excel"
                    elif "Word" in key_path:
                        app = "Word"
                    elif "PowerPoint" in key_path:
                        app = "PowerPoint"

                    found.append({
                        "name": subkey_name,
                        "reg_key": key_path,
                        "reg_subkey": subkey_name,
                        "manifest": manifest,
                        "app": app,
                        "source": "registry_subkey",
                    })
                except OSError:
                    pass
                i += 1
        except OSError:
            pass

        # Check for OPEN values (e.g. OPEN, OPEN1, OPEN2, ... in Excel Options)
        if "Options" in key_path:
            try:
                j = 0
                while True:
                    val_name, val_data, _ = winreg.EnumValue(key, j)
                    if val_name.upper().startswith("OPEN"):
                        app = "Office"
                        if "Excel" in key_path:
                            app = "Excel"
                        elif "Word" in key_path:
                            app = "Word"
                        elif "PowerPoint" in key_path:
                            app = "PowerPoint"

                        # val_data is typically the file path or /R "path"
                        display = val_data.strip().strip('"').strip("'")
                        if display.startswith("/R"):
                            display = display[2:].strip().strip('"')

                        found.append({
                            "name": f"{val_name} → {Path(display).name if display else val_data}",
                            "reg_key": key_path,
                            "reg_value_name": val_name,
                            "reg_value_data": val_data,
                            "app": app,
                            "source": "registry_value",
                        })
                    j += 1
            except OSError:
                pass

        winreg.CloseKey(key)

    return found


# ═════════════════════════════════════════════════════════════════════════════
# REMOVAL
# ═════════════════════════════════════════════════════════════════════════════

def remove_file(item: dict) -> tuple[bool, str]:
    """Delete an add-in file from disk."""
    try:
        item["path"].unlink()
        return True, f"Deleted file: {item['path']}"
    except PermissionError:
        return False, f"Cannot delete (is {item['app']} open?): {item['path']}"
    except Exception as e:
        return False, f"Error deleting {item['path']}: {e}"


def remove_registry_subkey(item: dict) -> tuple[bool, str]:
    """Delete a registry subkey (COM add-in)."""
    full_path = f"{item['reg_key']}\\{item['reg_subkey']}"
    try:
        # Delete all values in subkey first, then the subkey itself
        key = winreg.OpenKey(HIVE, item["reg_key"], 0, winreg.KEY_WRITE)
        winreg.DeleteKey(key, item["reg_subkey"])
        winreg.CloseKey(key)
        return True, f"Removed registry key: HKCU\\{full_path}"
    except FileNotFoundError:
        return True, f"Already gone: HKCU\\{full_path}"
    except PermissionError:
        return False, f"Permission denied: HKCU\\{full_path}"
    except Exception as e:
        return False, f"Error removing HKCU\\{full_path}: {e}"


def remove_registry_value(item: dict) -> tuple[bool, str]:
    """Delete a single registry value (OPEN, OPEN1, etc.)."""
    val_name = item["reg_value_name"]
    try:
        key = winreg.OpenKey(HIVE, item["reg_key"], 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(key, val_name)
        winreg.CloseKey(key)
        return True, f"Removed registry value: HKCU\\{item['reg_key']}  →  {val_name}"
    except FileNotFoundError:
        return True, f"Already gone: {val_name}"
    except PermissionError:
        return False, f"Permission denied: {val_name}"
    except Exception as e:
        return False, f"Error removing {val_name}: {e}"


def remove_item(item: dict) -> list[tuple[bool, str]]:
    """Remove an add-in (file and/or registry entry)."""
    results = []

    if item["source"] == "file":
        results.append(remove_file(item))
        # Also try to find and remove matching registry entries
        reg_items = scan_registry_addins()
        filename = item["name"].lower()
        for ri in reg_items:
            match = False
            if ri["source"] == "registry_value":
                if filename in ri.get("reg_value_data", "").lower():
                    match = True
            elif ri["source"] == "registry_subkey":
                if filename in ri.get("manifest", "").lower() or filename in ri["name"].lower():
                    match = True
            if match:
                if ri["source"] == "registry_subkey":
                    results.append(remove_registry_subkey(ri))
                elif ri["source"] == "registry_value":
                    results.append(remove_registry_value(ri))

    elif item["source"] == "registry_subkey":
        results.append(remove_registry_subkey(item))

    elif item["source"] == "registry_value":
        results.append(remove_registry_value(item))

    return results


# ═════════════════════════════════════════════════════════════════════════════
# PARSE USER SELECTION  (supports: "1", "1,3,5", "1-4", "1-3,5,7-9", "all")
# ═════════════════════════════════════════════════════════════════════════════

def parse_selection(text: str, max_val: int) -> list[int] | None:
    """Parse user input into a sorted list of 0-based indices."""
    text = text.strip().lower()
    if text in ("all", "a", "*"):
        return list(range(max_val))

    indices = set()
    for part in text.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            bounds = part.split("-", 1)
            try:
                lo, hi = int(bounds[0]), int(bounds[1])
            except ValueError:
                return None
            if lo > hi:
                lo, hi = hi, lo
            for v in range(lo, hi + 1):
                indices.add(v)
        else:
            try:
                indices.add(int(part))
            except ValueError:
                return None

    # Convert to 0-based and validate
    result = []
    for v in sorted(indices):
        idx = v - 1
        if idx < 0 or idx >= max_val:
            return None
        result.append(idx)
    return result


# ═════════════════════════════════════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════════════════════════════════════

def main():
    print()
    print("=" * 52)
    print("  Office Add-in Remover  (Clean Uninstall)")
    print("=" * 52)

    # Gather everything
    file_addins = scan_addin_files()
    reg_addins = scan_registry_addins()

    # Merge: files first, then registry-only entries (skip duplicates)
    all_items = list(file_addins)
    file_names = {a["name"].lower() for a in file_addins}

    for ri in reg_addins:
        # Skip registry entries that clearly match a file we already listed
        ri_id = ri["name"].lower()
        if ri["source"] == "registry_value":
            data = ri.get("reg_value_data", "").lower()
            if any(fn in data for fn in file_names):
                continue
        elif ri["source"] == "registry_subkey":
            manifest = ri.get("manifest", "").lower()
            if any(fn in manifest for fn in file_names) or ri_id in file_names:
                continue
        all_items.append(ri)

    if not all_items:
        print("\nNo add-ins found (files or registry).")
        input("\nPress Enter to exit...")
        return

    # Display
    print("\nFound add-ins:\n")
    for i, a in enumerate(all_items, 1):
        tag = a["app"]
        if a["source"].startswith("registry"):
            tag += " / registry"
        print(f"  [{i:>2}]  {a['name']}  ({tag})")

    print(f"\n  Total: {len(all_items)}")
    print()
    print("  Select items to remove:")
    print("    • Single:   3")
    print("    • Multiple: 1,3,5")
    print("    • Range:    1-4")
    print("    • Mixed:    1-3,5,7")
    print("    • All:      all")
    print()

    choice = input("  Your selection (or q to quit): ").strip()

    if choice.lower() in ("q", ""):
        print("  Cancelled.")
        input("\nPress Enter to exit...")
        return

    selected_indices = parse_selection(choice, len(all_items))
    if selected_indices is None:
        print("  Invalid selection.")
        input("\nPress Enter to exit...")
        return

    selected = [all_items[i] for i in selected_indices]

    print(f"\n  You selected {len(selected)} add-in(s):\n")
    for s in selected:
        loc = str(s.get("path", "")) or f"HKCU\\{s.get('reg_key', '')}\\{s.get('reg_subkey', s.get('reg_value_name', ''))}"
        print(f"    • {s['name']}  →  {loc}")

    print()
    confirm = input(f"  Delete these {len(selected)} add-in(s)? (y/n): ").strip().lower()

    if confirm != "y":
        print("  Cancelled.")
        input("\nPress Enter to exit...")
        return

    # Remove
    print()
    ok_count = 0
    err_count = 0
    for s in selected:
        results = remove_item(s)
        for success, msg in results:
            icon = "  ✓" if success else "  ✗"
            print(f"  {icon}  {msg}")
            if success:
                ok_count += 1
            else:
                err_count += 1

    print()
    print(f"  Done — {ok_count} action(s) succeeded, {err_count} failed.")

    if err_count > 0:
        print("\n  Tip: Close all Office apps and retry if you got permission errors.")

    print()
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()
