"""Self-contained MS-OVBA / OLE compound file helpers used by build_dotm.py.

Nothing here needs Word, PowerShell or COM: a .dotm is just an OOXML zip whose
word/vbaProject.bin is an OLE compound file holding the VBA project.  To
rebuild the template we patch the VBA module streams inside a seed
vbaProject.bin, so the recorded type library references stay byte identical to
the ones Word itself produced.
"""

import struct

# ---------------------------------------------------------------------------
# MS-OVBA compression (2.4.1.3.1).
#
# Per MS-OVBA 2.4.1.1.4, every CompressedChunk (except the last) MUST cover
# 4096 uncompressed source bytes.  Each chunk uses LZ77 CopyTokens per
# 2.4.1.3.19, falling back to an uncompressed raw chunk (CompressedChunkFlag=0)
# only if an incompressible 4096-byte block expands past 4094 body bytes.
# ---------------------------------------------------------------------------

_CHUNK_SIZE = 4096
_MAX_COMPRESSED_BODY = 4094  # header(2) + body <= 4096


def _compress_chunk(chunk):
    """Compress a single <=4096-byte chunk per MS-OVBA 2.4.1.3.1."""
    n = len(chunk)
    body = bytearray()
    pos = 0
    seen = {}  # 3-byte prefix -> list of offsets in chunk

    def index_byte(idx):
        if idx + 3 <= n:
            seen.setdefault(chunk[idx:idx + 3], []).append(idx)

    while pos < n:
        flag_pos = len(body)
        body.append(0x00)
        flag_byte = 0
        for bit in range(8):
            if pos >= n:
                break
            best_len = 0
            best_off = 0
            if pos > 0 and pos + 3 <= n:
                bit_count = max(4, (pos - 1).bit_length())
                max_len = min((0xFFFF >> bit_count) + 3, n - pos)
                max_off = min(1 << bit_count, pos)
                min_cand = pos - max_off
                cands = seen.get(chunk[pos:pos + 3])
                if cands:
                    for cand in reversed(cands):
                        if cand < min_cand:
                            break
                        if cand >= pos:
                            continue
                        mlen = 3
                        while mlen < max_len and chunk[cand + mlen] == chunk[pos + mlen]:
                            mlen += 1
                        if mlen > best_len:
                            best_len = mlen
                            best_off = pos - cand
                            if best_len == max_len:
                                break
            if best_len >= 3:
                bit_count = max(4, (pos - 1).bit_length())
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                flag_byte |= (1 << bit)
                body.extend(struct.pack("<H", token))
                for k in range(best_len):
                    index_byte(pos + k)
                pos += best_len
            else:
                body.append(chunk[pos])
                index_byte(pos)
                pos += 1
            if len(body) > _MAX_COMPRESSED_BODY:
                if n == _CHUNK_SIZE:
                    return struct.pack("<H", 0x3FFF) + chunk
                raise ValueError("Final MS-OVBA chunk body exceeded 4094 bytes")
        body[flag_pos] = flag_byte

    header = 0x8000 | 0x3000 | ((len(body) - 1) & 0x0FFF)
    return struct.pack("<H", header) + bytes(body)


def compress(data):
    """Wrap source bytes in an MS-OVBA container (4096-byte chunks)."""
    out = bytearray(b"\x01")
    if not data:
        return bytes(out)
    for pos in range(0, len(data), _CHUNK_SIZE):
        out.extend(_compress_chunk(data[pos:pos + _CHUNK_SIZE]))
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

def _module_offset_locations(dir_stream):
    """Return [(module_stream_name, text_offset_pos, text_offset)] from dir."""
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
        return []

    result = []
    # PROJECTCOOKIE
    rec_id, size = struct.unpack_from("<HI", dir_stream, pos)
    if rec_id != 0x0013:
        return []
    pos += 6 + size

    def sized(at):
        rid, sz = struct.unpack_from("<HI", dir_stream, at)
        return rid, sz, at + 6

    for _ in range(count):
        # MODULENAME
        rid, sz, cur = sized(pos)
        if rid != 0x0019:
            return result
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
        result.append((stream_name, cur, struct.unpack_from("<I", dir_stream, cur)[0]))
        pos = cur + sz
        # remaining per-module records up to the 0x002B terminator
        while pos + 6 <= n:
            rid, sz = struct.unpack_from("<HI", dir_stream, pos)
            if rid == 0x002B:
                pos += 6 + sz
                break
            pos += 6 + sz
    return result


def parse_module_offsets(dir_stream):
    """Return {module_stream_name: text_offset} from a decompressed dir."""
    return {name: off for name, _, off in _module_offset_locations(dir_stream)}


def zero_module_offsets(dir_stream):
    """Return a copy of decompressed dir with every MODULEOFFSET set to 0.

    When PerformanceCache is stripped from the module streams and _VBA_PROJECT
    is reset to uncompiled, TextOffset must be 0 so Word compiles cleanly from
    CompressedSourceCode instead of reading stale p-code.
    """
    buf = bytearray(dir_stream)
    locs = _module_offset_locations(dir_stream)
    if not locs:
        raise ValueError("Could not locate module offsets in dir stream")
    for _, cur, _ in locs:
        struct.pack_into("<I", buf, cur, 0)
    return bytes(buf)


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
                node.start = 0
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


def read_cfb(data):
    """Minimal OLE compound file reader -> {stream_path: bytes}."""
    if len(data) < 512 or data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise ValueError("not an OLE compound file")

    sector_shift, mini_shift = struct.unpack_from("<HH", data, 30)
    sec_sz = 1 << sector_shift
    mini_sz = 1 << mini_shift
    num_fat, dir_start = struct.unpack_from("<II", data, 44)
    mini_cutoff, minifat_start = struct.unpack_from("<II", data, 56)
    difat_start, difat_count = struct.unpack_from("<II", data, 68)

    def get_sec(idx):
        off = (idx + 1) * sec_sz
        return data[off:off + sec_sz]

    fat_sectors = [s for s in struct.unpack_from("<109I", data, 76) if s < 0xFFFFFFFC]
    cur_difat = difat_start
    for _ in range(difat_count):
        if cur_difat >= 0xFFFFFFFC:
            break
        buf = get_sec(cur_difat)
        ints = struct.unpack("<" + "I" * (sec_sz // 4), buf)
        fat_sectors.extend(s for s in ints[:-1] if s < 0xFFFFFFFC)
        cur_difat = ints[-1]

    fat = []
    for s in fat_sectors[:num_fat]:
        fat.extend(struct.unpack("<" + "I" * (sec_sz // 4), get_sec(s)))

    def read_chain(start):
        out = bytearray()
        s = start
        seen = set()
        while s < 0xFFFFFFFC:
            if s in seen:
                raise ValueError("cycle in CFB FAT chain")
            seen.add(s)
            out.extend(get_sec(s))
            s = fat[s]
        return bytes(out)

    dir_data = read_chain(dir_start)
    minifat = []
    if minifat_start < 0xFFFFFFFC:
        mf = read_chain(minifat_start)
        minifat = list(struct.unpack("<" + "I" * (len(mf) // 4), mf))

    entries = []
    for i in range(0, len(dir_data), 128):
        chunk = dir_data[i:i + 128]
        kind = chunk[66]
        if kind == 0:
            entries.append(None)
            continue
        nlen = struct.unpack_from("<H", chunk, 64)[0]
        name = chunk[:max(0, nlen - 2)].decode("utf-16-le", "replace")
        left, right, child = struct.unpack_from("<III", chunk, 68)
        start, size = struct.unpack_from("<IQ", chunk, 116)
        entries.append((name, kind, left, right, child, start, size))

    root_mini = b""
    root_child = NOSTREAM
    for entry in entries:
        if entry and entry[1] == 5:
            root_child = entry[4]
            if entry[5] < 0xFFFFFFFC and entry[6] > 0:
                root_mini = read_chain(entry[5])[:entry[6]]
            break

    def read_mini(start, size):
        out = bytearray()
        s = start
        seen = set()
        while s < 0xFFFFFFFC and len(out) < size:
            if s in seen:
                raise ValueError("cycle in CFB MiniFAT chain")
            seen.add(s)
            out.extend(root_mini[s * mini_sz:(s + 1) * mini_sz])
            s = minifat[s]
        return bytes(out[:size])

    streams = {}

    def walk(idx, prefix):
        if idx == NOSTREAM or idx >= len(entries) or entries[idx] is None:
            return
        name, kind, left, right, child, start, size = entries[idx]
        walk(left, prefix)
        path = (prefix + "/" + name) if prefix else name
        if kind == 2:
            if size == 0:
                streams[path] = b""
            elif size < mini_cutoff:
                streams[path] = read_mini(start, size)
            else:
                streams[path] = read_chain(start)[:size]
        elif kind == 1:
            walk(child, path)
        walk(right, prefix)

    walk(root_child, "")
    return streams