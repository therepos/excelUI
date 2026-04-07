"""
Purpose: Seal a VBA project in an Office add-in (.xlam/.dotm/.ppam).
Usage:
    1. python rel-vbaseal.py <input_file> [output_file]
    2. If output_file is omitted, the input file is overwritten.
Note: Dependency of ci-rel.yml
"""

import io
import os
import re
import sys
import zipfile

SEAL_ID  = '{00000000-0000-0000-0000-000000000000}'
SEAL_CMG = '00'
SEAL_DPB = '00'
SEAL_GC  = '00'

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


def modify_vba_bin(vba_data: bytes) -> bytes | None:
    """
    Modify the vbaProject.bin (OLE compound file) to inject seal values
    into the PROJECT stream, making it unviewable.
    """
    data = bytearray(vba_data)

    new_id  = f'ID="{SEAL_ID}"'.encode('latin-1')
    new_cmg = f'CMG="{SEAL_CMG}"'.encode('latin-1')
    new_dpb = f'DPB="{SEAL_DPB}"'.encode('latin-1')
    new_gc  = f'GC="{SEAL_GC}"'.encode('latin-1')

    # Find existing protection fields
    id_match  = re.search(rb'ID="[^"]*"', data)
    cmg_match = re.search(rb'CMG="[^"]*"', data)
    dpb_match = re.search(rb'DPB="[^"]*"', data)
    gc_match  = re.search(rb'GC="[^"]*"', data)

    if cmg_match or dpb_match or gc_match:
        # Existing protection found — replace values
        # Work backwards to preserve byte offsets
        replacements = []

        if id_match:
            replacements.append((id_match.start(), id_match.end(), new_id))
        if gc_match:
            replacements.append((gc_match.start(), gc_match.end(), new_gc))
        if dpb_match:
            replacements.append((dpb_match.start(), dpb_match.end(), new_dpb))
        if cmg_match:
            replacements.append((cmg_match.start(), cmg_match.end(), new_cmg))

        replacements.sort(key=lambda x: x[0], reverse=True)

        for start, end, new_val in replacements:
            old_len = end - start
            new_len = len(new_val)

            if new_len <= old_len:
                # Pad with spaces to maintain OLE sector size
                padded = new_val + b' ' * (old_len - new_len)
                data[start:end] = padded
            else:
                data[start:end] = new_val

    else:
        # No existing protection — insert CMG/DPB/GC lines
        # Find insertion point: before [Workspace] or after last Name= line
        workspace_match = re.search(rb'\[Workspace\]', data)
        name_match = re.search(rb'Name="[^"]*"\r\n', data)

        if workspace_match:
            insert_pos = workspace_match.start()
        elif name_match:
            insert_pos = name_match.end()
        else:
            print("  ERROR: Cannot find insertion point in PROJECT stream")
            return None

        insert_text = (
            f'ID="{SEAL_ID}"\r\n'
            f'CMG="{SEAL_CMG}"\r\n'
            f'DPB="{SEAL_DPB}"\r\n'
            f'GC="{SEAL_GC}"\r\n'
        ).encode('latin-1')

        data[insert_pos:insert_pos] = insert_text

    return bytes(data)


def seal_vba_project(input_path: str, output_path: str) -> bool:
    """
    Seal the VBA project inside an Office file by corrupting the
    password hash, making it unviewable in the VBA editor.
    """
    print(f"  Input:  {input_path}")
    print(f"  Output: {output_path}")

    with zipfile.ZipFile(input_path, 'r') as zf_in:
        vba_path = find_vbaproject_bin_path(zf_in)
        if not vba_path:
            print("  ERROR: No vbaProject.bin found in file")
            return False

        print(f"  Found:  {vba_path}")

        vba_data = zf_in.read(vba_path)

        if b'DPB="' in vba_data:
            print("  Note:   Existing protection found, replacing...")

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

    with open(output_path, 'wb') as f:
        f.write(buffer.getvalue())

    print("  Done!   VBA project sealed (unviewable).")
    return True


def main():
    if len(sys.argv) < 2:
        print("Usage: python rel-vbaseal.py <input_file> [output_file]")
        print()
        print("Seals the VBA project so it is unviewable in the editor.")
        print("No password is set — the project simply cannot be opened.")
        print("If output_file is omitted, input file is overwritten.")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else input_path

    if not os.path.exists(input_path):
        print(f"ERROR: File not found: {input_path}")
        sys.exit(1)

    print()
    print("=" * 40)
    print("  VBA Project Seal")
    print("=" * 40)

    success = seal_vba_project(input_path, output_path)

    if not success:
        sys.exit(1)

    print()


if __name__ == "__main__":
    main()
