"""Self-contained MS-OVBA / OLE compound file helpers used by build_dotm.py.

Nothing here needs Word, PowerShell or COM: a .dotm is just an OOXML zip whose
word/vbaProject.bin is an OLE compound file holding the VBA project.  To
rebuild the template we patch the VBA module streams inside a seed
vbaProject.bin, so the recorded type library references stay byte identical to
the ones Word itself produced.
"""

import struct

# ---------------------------------------------------------------------------
# MS-OVBA compression (2.4.1).  Only literal tokens are emitted: every group of
# up to 8 source bytes becomes a 0x00 flag byte followed by those bytes.  That
# is valid per spec and keeps the writer simple.
#
# Chunk sizing is deliberately conservative.  A literal-only chunk encodes n
# source bytes as n + ceil(n/8) body bytes, and the 12-bit CompressedChunkSize
# field tops out at 4095.  Packing to the theoretical maximum (3640 source
# bytes -> 4095 body bytes) produces a container that oletools decodes
# perfectly but that Word itself rejects: the template loads as an add-in and
# then exposes no macros at all.
#
# That was found by bisecting the builder against a real Word 16.89 install -
# source chunks of 2048, 1024 and 512 bytes are accepted, 3640 is not.  The
# writer therefore stays at 1024, far away from the boundary, since a literal
# encoding is ~12% larger than the source anyway and there is nothing to gain
# from packing tightly.
# ---------------------------------------------------------------------------

_LITERALS_PER_GROUP = 8
_MAX_SOURCE_PER_CHUNK = 1024
_MAX_BODY_PER_CHUNK = _MAX_SOURCE_PER_CHUNK + (_MAX_SOURCE_PER_CHUNK // 8) + 1


def compress(data):
    """Wrap source bytes in an MS-OVBA container."""
    out = bytearray(b"\x01")
    if not data:
        return bytes(out)
    pos = 0
    while pos < len(data):
        take = min(_MAX_SOURCE_PER_CHUNK, len(data) - pos)
        body = bytearray()
        group = 0
        while group < take:
            body.append(0x00)
            body.extend(data[pos + group:pos + group + _LITERALS_PER_GROUP])
            group += _LITERALS_PER_GROUP
        # CompressedChunkSize = len(body) + 2 (header), stored minus 3.
        assert len(body) <= _MAX_BODY_PER_CHUNK, "chunk body too large"
        assert (len(body) - 1) <= 0x0FFF, "CompressedChunkSize would overflow"
        header = 0x8000 | 0x3000 | ((len(body) - 1) & 0x0FFF)
        out.extend(struct.pack("<H", header))
        out.extend(body)
        pos += take
    return bytes(out)


def decompress(data):
    """Inverse of compress()."""
    if not data or data[0] != 0x01:
        raise ValueError("not an MS-OVBA container")
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        compressed = bool(header & 0x8000)
        body = data[pos + 2:pos + size]
        pos += size
        chunk_origin = len(out)
        if not compressed:
            out.extend(body)
            continue
        i = 0
        while i < len(body):
            flags = body[i]
            i += 1
            for bit in range(8):
                if i >= len(body):
                    break
                if not (flags >> bit) & 1:
                    out.append(body[i])
                    i += 1
                    continue
                token = struct.unpack_from("<H", body, i)[0]
                i += 2
                diff = len(out) - chunk_origin
                bit_count = max(4, (diff - 1).bit_length())
                length_mask = 0xFFFF >> bit_count
                offset = ((token & ~length_mask) >> (16 - bit_count)) + 1
                length = (token & length_mask) + 3
                for _ in range(length):
                    out.append(out[len(out) - offset])
    return bytes(out)


# ---------------------------------------------------------------------------
# dir stream: we only need each module's TextOffset so the new source can be
# written at exactly the same place inside the module stream.
# ---------------------------------------------------------------------------

def parse_module_offsets(dir_stream):
    """Return {module_stream_name: text_offset} from a decompressed dir.

    Everything before PROJECTMODULES is skipped by scanning for record 0x000F,
    which is cheap and safe: the type library references in between have
    variable-length records that we do not need to understand.  From there on
    the layout is fixed by MS-OVBA 2.3.4.2 and is parsed precisely.
    """
    n = len(dir_stream)
    pos = -1
    for i in range(n - 6):
        if struct.unpack_from("<H", dir_stream, i)[0] != 0x000F:
            continue
        if struct.unpack_from("<I", dir_stream, i + 2)[0] != 2:
            continue
        count = struct.unpack_from("<H", dir_stream, i + 6)[0]
        pos = i + 8
        break
    if pos < 0:
        return {}

    result = {}
    # PROJECTCOOKIE
    rec_id, size = struct.unpack_from("<HI", dir_stream, pos)
    if rec_id != 0x0013:
        return {}
    pos += 6 + size

    def sized(at):
        rid, sz = struct.unpack_from("<HI", dir_stream, at)
        return rid, sz, at + 6

    for _ in range(count):
        # MODULENAME
        rid, sz, cur = sized(pos)
        if rid != 0x0019:
            return result
        module_name = dir_stream[cur:cur + sz].decode("latin-1")
        pos = cur + sz
        # MODULENAMEUNICODE
        rid, sz, cur = sized(pos)
        if rid != 0x0047:
            return result
        pos = cur + sz
        # MODULESTREAMNAME (+ reserved unicode copy)
        rid, sz, cur = sized(pos)
        if rid != 0x001A:
            return result
        stream_name = dir_stream[cur:cur + sz].decode("latin-1")
        pos = cur + sz
        rid, sz, cur = sized(pos)          # 0x0032 reserved
        pos = cur + sz
        # MODULEDOCSTRING (+ reserved unicode copy)
        rid, sz, cur = sized(pos)
        if rid != 0x001C:
            return result
        pos = cur + sz
        rid, sz, cur = sized(pos)          # 0x0048 reserved
        pos = cur + sz
        # MODULEOFFSET
        rid, sz, cur = sized(pos)
        if rid != 0x0031:
            return result
        result[stream_name] = struct.unpack_from("<I", dir_stream, cur)[0]
        pos = cur + sz
        # remaining per-module records up to the 0x002B terminator
        while pos + 6 <= n:
            rid, sz = struct.unpack_from("<HI", dir_stream, pos)
            if rid == 0x002B:
                pos += 6 + sz
                break
            pos += 6 + sz
        module_name = module_name
    return result


# ---------------------------------------------------------------------------
# Minimal OLE compound file writer (512 byte sectors, version 3)
# ---------------------------------------------------------------------------

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

SECTOR = 512
MINI_SECTOR = 64
MINI_CUTOFF = 4096


class _Entry:
    def __init__(self, name, kind, data=b""):
        self.name = name
        self.kind = kind                  # 5 root, 1 storage, 2 stream
        self.data = data
        self.child = NOSTREAM
        self.left = NOSTREAM
        self.right = NOSTREAM
        self.start = ENDOFCHAIN
        self.size = 0
        self.index = -1
        self.parent = None


def _sort_key(entry):
    # CFB orders directory entries by name length first, then uppercase name.
    return (len(entry.name), entry.name.upper())


def _build_tree(entries):
    """Balanced BST over CFB ordering; all nodes marked black."""
    if not entries:
        return NOSTREAM
    ordered = sorted(entries, key=_sort_key)

    def build(lo, hi):
        if lo > hi:
            return NOSTREAM
        mid = (lo + hi) // 2
        node = ordered[mid]
        node.left = build(lo, mid - 1)
        node.right = build(mid + 1, hi)
        return node.index

    return build(0, len(ordered) - 1)


def write_cfb(tree):
    """tree: {"VBA": {"dir": b"...", "Mod": b"..."}, "PROJECT": b"..."}.

    Returns the compound file bytes.
    """
    root = _Entry("Root Entry", 5)
    root.index = 0
    entries = [root]
    storages = {}

    def add(parent, mapping):
        kids = []
        for name, value in mapping.items():
            if isinstance(value, dict):
                node = _Entry(name, 1)
                node.index = len(entries)
                node.parent = parent
                entries.append(node)
                storages[id(node)] = node
                add(node, value)
            else:
                node = _Entry(name, 2, value)
                node.index = len(entries)
                node.parent = parent
                entries.append(node)
            kids.append(node)
        parent.child = _build_tree(kids)

    add(root, tree)

    mini_streams = [e for e in entries
                    if e.kind == 2 and len(e.data) < MINI_CUTOFF]
    big_streams = [e for e in entries
                   if e.kind == 2 and len(e.data) >= MINI_CUTOFF]

    # Mini stream + mini FAT
    mini_data = bytearray()
    minifat = []
    for e in mini_streams:
        first = len(mini_data) // MINI_SECTOR
        padded = e.data + b"\x00" * (-len(e.data) % MINI_SECTOR)
        count = len(padded) // MINI_SECTOR
        mini_data.extend(padded)
        for k in range(count):
            minifat.append(first + k + 1 if k < count - 1 else ENDOFCHAIN)
        e.start = first
        e.size = len(e.data)
    for e in big_streams:
        e.size = len(e.data)
    root.size = len(mini_data)

    minifat_bytes = b"".join(struct.pack("<I", v) for v in minifat)
    if minifat_bytes:
        pad = -len(minifat_bytes) % SECTOR
        minifat_bytes += struct.pack("<I", FREESECT) * (pad // 4)

    dir_bytes_len = ((len(entries) * 128 + SECTOR - 1) // SECTOR) * SECTOR

    sectors = []
    fat = []

    def alloc(payload):
        if not payload:
            return ENDOFCHAIN
        first = len(sectors)
        blocks = [payload[i:i + SECTOR].ljust(SECTOR, b"\x00")
                  for i in range(0, len(payload), SECTOR)]
        for i, b in enumerate(blocks):
            sectors.append(b)
            fat.append(first + i + 1 if i < len(blocks) - 1 else ENDOFCHAIN)
        return first

    for e in big_streams:
        e.start = alloc(e.data)

    root.start = alloc(bytes(mini_data))
    minifat_start = alloc(minifat_bytes)
    minifat_count = len(minifat_bytes) // SECTOR if minifat_bytes else 0

    dir_payload = bytearray(b"\x00" * dir_bytes_len)
    for e in entries:
        off = e.index * 128
        raw = e.name.encode("utf-16-le") + b"\x00\x00"
        dir_payload[off:off + len(raw)] = raw
        struct.pack_into("<H", dir_payload, off + 64, len(raw))
        dir_payload[off + 66] = e.kind
        dir_payload[off + 67] = 1                    # black
        struct.pack_into("<III", dir_payload, off + 68,
                         e.left, e.right, e.child)
        struct.pack_into("<I", dir_payload, off + 116, e.start)
        struct.pack_into("<Q", dir_payload, off + 120, e.size)
    dir_start = alloc(bytes(dir_payload))

    # FAT sectors describe themselves, so solve for the count.
    n_fat = 1
    while True:
        total = len(sectors) + n_fat
        if (total * 4 + SECTOR - 1) // SECTOR <= n_fat:
            break
        n_fat += 1
    fat_first = len(sectors)
    for _ in range(n_fat):
        sectors.append(b"\x00" * SECTOR)
        fat.append(FATSECT)

    fat_entries = fat + [FREESECT] * (n_fat * (SECTOR // 4) - len(fat))
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat_entries)
    for i in range(n_fat):
        sectors[fat_first + i] = fat_bytes[i * SECTOR:(i + 1) * SECTOR]

    header = bytearray(b"\x00" * SECTOR)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<HHH", header, 24, 0x003E, 0x0003, 0xFFFE)
    struct.pack_into("<HH", header, 30, 9, 6)
    struct.pack_into("<I", header, 40, 0)              # dir sector count (v3)
    struct.pack_into("<I", header, 44, n_fat)
    struct.pack_into("<I", header, 48, dir_start)
    struct.pack_into("<I", header, 52, 0)
    struct.pack_into("<I", header, 56, MINI_CUTOFF)
    struct.pack_into("<I", header, 60, minifat_start)
    struct.pack_into("<I", header, 64, minifat_count)
    struct.pack_into("<I", header, 68, ENDOFCHAIN)     # first DIFAT sector
    struct.pack_into("<I", header, 72, 0)              # DIFAT sector count
    for i in range(109):
        struct.pack_into("<I", header, 76 + i * 4,
                         fat_first + i if i < n_fat else FREESECT)

    return bytes(header) + b"".join(sectors)