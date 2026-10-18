"""
Purpose: Build an Office add-in (.xlam) from VBA source files - no Excel needed.
Usage:
    python vbabuild.py <src_dir> <xlam>
    - <xlam> is used as the shell (sheets, styles, references) and is overwritten.
    - Every *.bas / *.cls in <src_dir> becomes a VBA module (file order = project order,
      ThisWorkbook/Sheet* document modules first).
    - <src_dir>/<name>.xml (if present) replaces customUI/customUI14.xml.
Note: Writes a source-only VBA project (no compiled p-code). Office recompiles the
      source on first open, as it does for files saved by LibreOffice. Dependency of ci-excelUI.yml
"""

import io
import os
import re
import struct
import sys
import zipfile

import olefile
from oletools.olevba import decompress_stream

CODEPAGE = "cp1252"

# ═════════════════════════════════════════════════════════════════════════════
# MS-OVBA compression (2.4.1)


def compress(data: bytes) -> bytes:
    out = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        body = bytearray()
        pos = 0
        while pos < len(chunk):
            flags, tokens = 0, bytearray()
            for bit in range(8):
                if pos >= len(chunk):
                    break
                # Bit widths for copy tokens depend on position within the chunk
                bits = max((pos - 1).bit_length(), 4) if pos > 0 else 4
                max_len = (0xFFFF >> bits) + 3
                max_off = 1 << bits
                best_len, best_off = 0, 0
                for cand in range(max(0, pos - max_off), pos):
                    ln = 0
                    while (ln < max_len and pos + ln < len(chunk)
                           and chunk[cand + ln] == chunk[pos + ln]):
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, pos - cand
                if best_len >= 3:
                    token = ((best_off - 1) << (16 - bits)) | (best_len - 3)
                    tokens += struct.pack("<H", token)
                    flags |= 1 << bit
                    pos += best_len
                else:
                    tokens.append(chunk[pos])
                    pos += 1
            body.append(flags)
            body += tokens
        if len(body) < 4096:
            out += struct.pack("<H", 0xB000 | (len(body) + 2 - 3))
            out += body
        else:  # Incompressible: raw chunk (must be exactly 4096 bytes)
            out += struct.pack("<H", 0x3000 | (4096 + 2 - 3))
            out += chunk.ljust(4096, b"\x00")
    return bytes(out)


# ═════════════════════════════════════════════════════════════════════════════
# Minimal Compound File (CFB v3) writer


class _Entry:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data  # kind: 1 storage, 2 stream, 5 root
        self.children = []
        self.left = self.right = self.child = 0xFFFFFFFF
        self.start, self.size = 0xFFFFFFFE, len(data)


def _key(e):
    return (len(e.name), e.name.upper())


def write_cfb(streams: dict) -> bytes:
    """streams: {'PROJECT': bytes, 'VBA/dir': bytes, ...} -> CFB file bytes."""
    root = _Entry("Root Entry", 5)
    storages = {"": root}
    for path, data in streams.items():
        parts = path.split("/")
        parent = root
        for i, p in enumerate(parts[:-1]):
            key = "/".join(parts[:i + 1])
            if key not in storages:
                storages[key] = _Entry(p, 1)
                parent.children.append(storages[key])
            parent = storages[key]
        parent.children.append(_Entry(parts[-1], 2, data))

    entries = []

    def flatten(e):
        entries.append(e)
        for c in e.children:
            flatten(c)
    flatten(root)
    idx = {id(e): i for i, e in enumerate(entries)}

    def tree(items):  # Balanced BST of siblings (all nodes black)
        if not items:
            return 0xFFFFFFFF
        mid = len(items) // 2
        node = items[mid]
        node.left, node.right = tree(items[:mid]), tree(items[mid + 1:])
        return idx[id(node)]
    for e in entries:
        e.child = tree(sorted(e.children, key=_key))

    # Mini stream for small streams
    mini = bytearray()
    for e in entries:
        if e.kind == 2 and e.size < 4096:
            e.start = len(mini) // 64
            mini += e.data + b"\x00" * (-len(e.data) % 64)
    minifat = []
    for e in entries:
        if e.kind == 2 and e.size < 4096:
            n = (e.size + 63) // 64
            minifat += [e.start + i + 1 for i in range(n - 1)] + [0xFFFFFFFE] if n else []
    root.data, root.size = bytes(mini), len(mini)

    sectors, fat = [], []

    def alloc(data):
        if not data:
            return 0xFFFFFFFE
        start = len(sectors)
        n = (len(data) + 511) // 512
        for i in range(n):
            sectors.append(data[i * 512:(i + 1) * 512].ljust(512, b"\x00"))
            fat.append(start + i + 1 if i < n - 1 else 0xFFFFFFFE)
        return start

    for e in entries:
        if e.kind == 2 and e.size >= 4096:
            e.start = alloc(e.data)
    root.start = alloc(root.data)
    minifat += [0xFFFFFFFF] * (-len(minifat) % 128)
    minifat_bytes = b"".join(struct.pack("<I", x) for x in minifat)
    minifat_start = alloc(minifat_bytes)

    direc = bytearray()
    for e in entries:
        name = e.name.encode("utf-16-le") + b"\x00\x00"
        d = name.ljust(64, b"\x00") + struct.pack("<HBB", len(name), e.kind, 1)
        d += struct.pack("<III", e.left, e.right, e.child)
        d += b"\x00" * 16 + b"\x00" * 4 + b"\x00" * 16
        d += struct.pack("<IQ", e.start if (e.kind == 2 or e.kind == 5) else 0, e.size if e.kind != 1 else 0)
        direc += d
    direc += b"\x00" * (-len(direc) % 512)
    dir_start = alloc(bytes(direc))

    # FAT sectors (each covers 128 entries, including themselves)
    nfat = 1
    while (len(sectors) + nfat) > nfat * 128:
        nfat += 1
    fat_start = len(sectors)
    fat += [0xFFFFFFFD] * nfat
    fat += [0xFFFFFFFF] * (-len(fat) % 128)
    if nfat > 109:
        raise ValueError("File too large for this writer")
    for i in range(nfat):
        sectors.append(b"".join(struct.pack("<I", x) for x in fat[i * 128:(i + 1) * 128]))

    hdr = bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
    hdr += struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\x00" * 6
    hdr += struct.pack("<IIIIIIIII", 0, nfat, dir_start, 0, 4096,
                       minifat_start, (len(minifat_bytes) + 511) // 512 if minifat else 0,
                       0xFFFFFFFE, 0)
    difat = [fat_start + i for i in range(nfat)] + [0xFFFFFFFF] * (109 - nfat)
    hdr += b"".join(struct.pack("<I", x) for x in difat)
    return hdr + b"".join(sectors)


# ═════════════════════════════════════════════════════════════════════════════
# VBA project


def _rec(rid, body: bytes) -> bytes:
    return struct.pack("<HI", rid, len(body)) + body


def load_modules(src_dir):
    mods = []
    for fn in sorted(os.listdir(src_dir)):
        base, ext = os.path.splitext(fn)
        if ext.lower() not in (".bas", ".cls"):
            continue
        with open(os.path.join(src_dir, fn), "rb") as f:
            text = f.read().decode("utf-8-sig").replace("\r\n", "\n").replace("\n", "\r\n")
        m = re.search(r'^Attribute VB_Name = "([^"]+)"', text, re.M)
        name = m.group(1) if m else re.sub(r"\W", "", base)
        if not m:
            text = f'Attribute VB_Name = "{name}"\r\n' + text
        doc = ext.lower() == ".cls" and "VB_Base" in text
        mods.append((name, doc, text))
    # Document modules first, keep ThisWorkbook on top
    mods.sort(key=lambda m: (not m[1], m[0] != "ThisWorkbook"))
    return mods


def template_dir_head(dir_bytes: bytes) -> bytes:
    """Return PROJECTINFORMATION + PROJECTREFERENCES records (everything before modules)."""
    i = 0
    while i < len(dir_bytes):
        rid, sz = struct.unpack_from("<HI", dir_bytes, i)
        if rid == 0x0F:
            return dir_bytes[:i]
        i += 6 + (6 if rid == 0x09 else sz)
    raise ValueError("PROJECTMODULES record not found in template")


def build_vba(template_bin: bytes, mods) -> bytes:
    ole = olefile.OleFileIO(template_bin)
    dir_head = template_dir_head(bytes(decompress_stream(bytearray(ole.openstream("VBA/dir").read()))))
    project_txt = ole.openstream("PROJECT").read().decode(CODEPAGE)
    ole.close()

    d = bytearray(dir_head)
    d += _rec(0x0F, struct.pack("<H", len(mods)))
    d += _rec(0x13, struct.pack("<H", 0xFFFF))
    streams = {}
    for name, doc, text in mods:
        a, u = name.encode(CODEPAGE), name.encode("utf-16-le")
        d += _rec(0x19, a) + _rec(0x47, u) + _rec(0x1A, a) + _rec(0x32, u)
        d += _rec(0x1C, b"") + _rec(0x48, b"")
        d += _rec(0x31, struct.pack("<I", 0))
        d += _rec(0x1E, struct.pack("<I", 0))
        d += _rec(0x2C, struct.pack("<H", 0xFFFF))
        d += _rec(0x22 if doc else 0x21, b"")
        if re.search(r"^Option Private Module", text, re.M | re.I):
            d += _rec(0x28, b"")
        d += _rec(0x2B, b"")
        streams["VBA/" + name] = compress(text.encode(CODEPAGE))
    d += struct.pack("<HI", 0x10, 0)

    # PROJECT stream: keep ID/protection/host lines from template, regenerate module lines
    keep = [ln for ln in project_txt.split("\r\n")
            if ln and not re.match(r"^(Document|Module|Class|BaseClass|Package)=", ln)]
    if "[Workspace]" in keep:
        keep = keep[:keep.index("[Workspace]")]
    lines = [ln for ln in keep if ln.startswith("ID=")]
    for name, doc, _ in mods:
        lines.append(f"Document={name}/&H00000000" if doc else f"Module={name}")
    for ln in keep:
        if not ln.startswith("ID="):
            lines += ["", ln] if ln.startswith("[") else [ln]
    lines += ["", "[Workspace]"] + [f"{n}=0, 0, 0, 0, C" for n, _, _ in mods]
    project = "\r\n".join(lines) + "\r\n"

    wm = b"".join(n.encode(CODEPAGE) + b"\x00" + n.encode("utf-16-le") + b"\x00\x00"
                  for n, _, _ in mods) + b"\x00\x00"

    streams.update({
        "PROJECT": project.encode(CODEPAGE),
        "PROJECTwm": wm,
        "VBA/_VBA_PROJECT": bytes.fromhex("CC61FFFF000300"),
        "VBA/dir": compress(bytes(d)),
    })
    return write_cfb(streams)


# ═════════════════════════════════════════════════════════════════════════════


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    src_dir, xlam = sys.argv[1], sys.argv[2]
    mods = load_modules(src_dir)
    xml_path = os.path.join(src_dir, os.path.splitext(os.path.basename(xlam))[0] + ".xml")

    with zipfile.ZipFile(xlam) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == "xl/vbaProject.bin":
                data = build_vba(data, mods)
            elif info.filename == "customUI/customUI14.xml" and os.path.exists(xml_path):
                with open(xml_path, "rb") as f:
                    data = f.read()
            zout.writestr(info, data)
    with open(xlam, "wb") as f:
        f.write(out.getvalue())
    print(f"Built {xlam}: " + ", ".join(n for n, _, _ in mods))


if __name__ == "__main__":
    main()
