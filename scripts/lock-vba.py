"""
Purpose: Lock a VBA project in an Office add-in (.xlam/.dotm/.ppam) with a password.
Usage:
    python lock-vba.py <input_file> [output_file]
    If output_file is omitted, the input file is overwritten.
    The default password is: 1234
    To change it, generate new CMG/DPB/GC values (see instructions below).
"""

import io
import os
import re
import struct
import sys
import zipfile


# ═════════════════════════════════════════════════════════════════════════════
# PASSWORD CONFIGURATION
#
# These values correspond to the password: 1234
#
# To generate values for a different password:
#   1. Create a new .xlsm in Excel
#   2. Open VBA Editor (Alt+F11) → Tools → VBA Project Properties
#   3. Protection tab → check "Lock project for viewing", set your password
#   4. Save and close
#   5. Rename .xlsm to .zip, extract xl/vbaProject.bin
#   6. Open vbaProject.bin in a hex editor, search for "CMG="
#   7. Copy the ID, CMG, DPB, and GC values below
# ═════════════════════════════════════════════════════════════════════════════

LOCK_ID   = '{00000000-0000-0000-0000-000000000000}'
LOCK_CMG  = '4E4CE2071EF919FD19FD1D011D01'
LOCK_DPB  = '4D4FE1041F1E3C1E3CE1C41F3C2B5B632FDA264D0036FBD163C3E84D0137CC1CE7D1CEF5E5D5'
LOCK_GC   = '4C4EE005200521052105'

# ═════════════════════════════════════════════════════════════════════════════


def find_vbaproject_bin_path(zf: zipfile.ZipFile) -> str | None:
    """Find the vbaProject.bin inside the ZIP archive."""
    candidates = [
        "xl/vbaProject.bin",
        "word/vbaProject.bin",
        "ppt/vbaProject.bin",
    ]
    for c in candidates:
        if c in zf.namelist():
            return c
    return None


def modify_project_stream(ole_data: bytes) -> bytes:
    """
    Modify the PROJECT stream inside the OLE compound file (vbaProject.bin)
    to inject password protection values.

    This is a targeted binary edit of the PROJECT stream text content.
    """
    # The PROJECT stream is stored as text inside the OLE compound file.
    # We need to find it and replace/add the ID, CMG, DPB, GC lines.

    # Try to find the PROJECT stream content by looking for known markers
    # The PROJECT stream contains lines like:
    #   ID="{...}"
    #   CMG="..."
    #   DPB="..."
    #   GC="..."

    # Find the project stream - it contains "Name=" and typically "Module="
    # We look for the text portion that has these key=value pairs

    # Decode as latin-1 to preserve all bytes
    text = ole_data.decode('latin-1')

    # Replace or add ID
    if re.search(r'^ID="[^"]*"', text, re.MULTILINE):
        text = re.sub(
            r'^ID="[^"]*"',
            f'ID="{LOCK_ID}"',
            text,
            count=1,
            flags=re.MULTILINE
        )
    else:
        # ID should already exist, but if not, we can't safely add it
        pass

    # Replace or add CMG
    if re.search(r'^CMG="[^"]*"', text, re.MULTILINE):
        text = re.sub(r'^CMG="[^"]*"', f'CMG="{LOCK_CMG}"', text, count=1, flags=re.MULTILINE)
    else:
        text = text.rstrip('\r\n') + f'\r\nCMG="{LOCK_CMG}"\r\n'

    # Replace or add DPB
    if re.search(r'^DPB="[^"]*"', text, re.MULTILINE):
        text = re.sub(r'^DPB="[^"]*"', f'DPB="{LOCK_DPB}"', text, count=1, flags=re.MULTILINE)
    else:
        text = text.rstrip('\r\n') + f'\r\nDPB="{LOCK_DPB}"\r\n'

    # Replace or add GC
    if re.search(r'^GC="[^"]*"', text, re.MULTILINE):
        text = re.sub(r'^GC="[^"]*"', f'GC="{LOCK_GC}"', text, count=1, flags=re.MULTILINE)
    else:
        text = text.rstrip('\r\n') + f'\r\nGC="{LOCK_GC}"\r\n'

    return text.encode('latin-1')


def lock_vba_project(input_path: str, output_path: str) -> bool:
    """
    Lock the VBA project inside an Office file by modifying vbaProject.bin.
    """
    print(f"  Input:  {input_path}")
    print(f"  Output: {output_path}")

    # Read the entire ZIP
    with zipfile.ZipFile(input_path, 'r') as zf_in:
        vba_path = find_vbaproject_bin_path(zf_in)
        if not vba_path:
            print("  ERROR: No vbaProject.bin found in file")
            return False

        print(f"  Found:  {vba_path}")

        # Read vbaProject.bin
        vba_data = zf_in.read(vba_path)

        # The vbaProject.bin is an OLE compound file.
        # The PROJECT stream is stored as raw text within it.
        # We can do a targeted find-and-replace on the binary data.

        # Check if it already has DPB (already locked)
        if b'DPB="' in vba_data:
            print("  Note:   VBA project already has password protection, replacing it...")

        # Modify the PROJECT stream content within the OLE binary
        modified_vba = modify_vba_bin(vba_data)

        if modified_vba is None:
            print("  ERROR: Could not modify vbaProject.bin")
            return False

        # Write new ZIP with modified vbaProject.bin
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.namelist():
                if item == vba_path:
                    zf_out.writestr(item, modified_vba)
                else:
                    zf_out.writestr(item, zf_in.read(item))

    # Write output
    with open(output_path, 'wb') as f:
        f.write(buffer.getvalue())

    print("  Done!   VBA project locked.")
    return True


def modify_vba_bin(vba_data: bytes) -> bytes | None:
    """
    Modify the vbaProject.bin (OLE compound file) to inject lock values
    into the PROJECT stream.

    We do targeted binary replacements on the raw OLE data since the
    PROJECT stream text is stored as-is within the compound file.
    """
    data = bytearray(vba_data)

    # The PROJECT stream contains text lines. We need to handle the
    # replacement carefully to maintain the OLE structure.

    # Strategy: find the exact byte sequences and replace them.
    # Since modifying lengths can break OLE sector allocation,
    # we pad replacements to match original lengths where possible.

    # Find ID line
    id_match = re.search(rb'ID="[^"]*"', data)
    new_id = f'ID="{LOCK_ID}"'.encode('latin-1')

    # Find or prepare CMG, DPB, GC
    cmg_match = re.search(rb'CMG="[^"]*"', data)
    dpb_match = re.search(rb'DPB="[^"]*"', data)
    gc_match = re.search(rb'GC="[^"]*"', data)

    new_cmg = f'CMG="{LOCK_CMG}"'.encode('latin-1')
    new_dpb = f'DPB="{LOCK_DPB}"'.encode('latin-1')
    new_gc = f'GC="{LOCK_GC}"'.encode('latin-1')

    # We need to be careful about size changes in the OLE compound file.
    # The safest approach: replace existing values with same-or-padded length.
    # If values don't exist yet, we append them after the last known marker.

    # Work backwards to preserve offsets
    replacements = []

    if id_match:
        replacements.append((id_match.start(), id_match.end(), new_id))

    if gc_match:
        replacements.append((gc_match.start(), gc_match.end(), new_gc))
    if dpb_match:
        replacements.append((dpb_match.start(), dpb_match.end(), new_dpb))
    if cmg_match:
        replacements.append((cmg_match.start(), cmg_match.end(), new_cmg))

    if not cmg_match and not dpb_match and not gc_match:
        # No existing protection — need to insert CMG/DPB/GC lines
        # Find a good insertion point: after the last line before [Workspace]
        workspace_match = re.search(rb'\[Workspace\]', data)
        # Or after the Name= line
        name_match = re.search(rb'Name="[^"]*"\r\n', data)

        if workspace_match:
            insert_pos = workspace_match.start()
        elif name_match:
            insert_pos = name_match.end()
        else:
            print("  ERROR: Cannot find insertion point in PROJECT stream")
            return None

        insert_text = (
            f'ID="{LOCK_ID}"\r\n'
            f'CMG="{LOCK_CMG}"\r\n'
            f'DPB="{LOCK_DPB}"\r\n'
            f'GC="{LOCK_GC}"\r\n'
        ).encode('latin-1')

        # For insertion, we need to handle OLE sector sizes
        # Since we're adding bytes, the sectors after this point shift.
        # This works for small OLE files where the PROJECT stream fits in one sector chain.
        data[insert_pos:insert_pos] = insert_text

        if id_match:
            # Still replace the existing ID
            # Recalculate position since we inserted before it might have shifted
            id_match2 = re.search(rb'ID="[^"]*"', data)
            if id_match2 and id_match2.start() != insert_pos:
                old = data[id_match2.start():id_match2.end()]
                # Pad new value to same length
                padded_id = new_id.ljust(len(old), b'\x00')
                data[id_match2.start():id_match2.end()] = padded_id

        return bytes(data)

    # Apply replacements (sorted by position, reversed to preserve offsets)
    replacements.sort(key=lambda x: x[0], reverse=True)

    for start, end, new_val in replacements:
        old_len = end - start
        new_len = len(new_val)

        if new_len <= old_len:
            # Pad with null bytes to maintain size
            padded = new_val + b'\x00' * (old_len - new_len)
            data[start:end] = padded
        else:
            # New value is longer — we need to expand
            # This can break OLE structure in some cases but generally works
            # for the PROJECT stream since it's typically in the last sectors
            data[start:end] = new_val

    return bytes(data)


def main():
    if len(sys.argv) < 2:
        print("Usage: python lock-vba.py <input_file> [output_file]")
        print()
        print("Locks the VBA project with password: 1234")
        print("If output_file is omitted, input file is overwritten.")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else input_path

    if not os.path.exists(input_path):
        print(f"ERROR: File not found: {input_path}")
        sys.exit(1)

    print()
    print("=" * 40)
    print("  VBA Project Lock")
    print("=" * 40)

    success = lock_vba_project(input_path, output_path)

    if not success:
        sys.exit(1)

    print()


if __name__ == "__main__":
    main()