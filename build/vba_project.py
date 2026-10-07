"""Génère un fichier vbaProject.bin (format MS-OVBA) à partir de sources VBA.

Le projet est écrit « source seule » : le flux _VBA_PROJECT porte la version
0xFFFF, ce qui oblige Excel à ignorer le cache p-code et à recompiler le code
source à l'ouverture. Aucun Excel n'est nécessaire pour produire le fichier.

Références : [MS-OVBA] (compression 2.4.1, flux dir 2.3.4.2, PROJECT 2.3.1,
chiffrement des données 2.4.3) et [MS-CFB] (fichier composé).
"""

import random
import struct
import uuid

CODEPAGE = 1252

# --------------------------------------------------------------------------
# Compression MS-OVBA (2.4.1)
# --------------------------------------------------------------------------


def _copy_token_help(difference):
    bit_count = 4
    while (1 << bit_count) < difference:
        bit_count += 1
    length_mask = 0xFFFF >> bit_count
    maximum_length = length_mask + 3
    return bit_count, maximum_length


def _compress_chunk(data):
    """Compresse au plus 4096 octets. Renvoie l'en-tête + les données."""
    out = bytearray()
    pos = 0
    end = len(data)
    index = {}  # trigramme -> positions précédentes dans le bloc

    def add_index(p):
        if p + 3 <= end:
            index.setdefault(bytes(data[p:p + 3]), []).append(p)

    while pos < end:
        flag_pos = len(out)
        out.append(0)
        flag = 0
        for bit in range(8):
            if pos >= end:
                break
            best_len, best_off = 0, 0
            if pos >= 1 and pos + 3 <= end:
                _, max_len = _copy_token_help(pos)
                for cand in reversed(index.get(bytes(data[pos:pos + 3]), [])[-64:]):
                    length = 0
                    limit = min(max_len, end - pos)
                    while length < limit and data[cand + length] == data[pos + length]:
                        length += 1
                    if length > best_len:
                        best_len, best_off = length, pos - cand
                        if length == limit:
                            break
            if best_len >= 3:
                bit_count, _ = _copy_token_help(pos)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flag |= 1 << bit
                for p in range(pos, pos + best_len):
                    add_index(p)
                pos += best_len
            else:
                out.append(data[pos])
                add_index(pos)
                pos += 1
        out[flag_pos] = flag

    if len(out) > 4096:
        # Bloc non compressé : exactement 4096 octets requis.
        assert len(data) == 4096, "bloc brut incomplet"
        header = 0x3000 | 0x0FFF
        return struct.pack("<H", header) + bytes(data)
    header = 0x8000 | 0x3000 | ((len(out) + 2 - 3) & 0x0FFF)
    return struct.pack("<H", header) + bytes(out)


def compress(data):
    out = bytearray(b"\x01")
    for i in range(0, len(data), 4096):
        out += _compress_chunk(data[i:i + 4096])
    return bytes(out)


# --------------------------------------------------------------------------
# Chiffrement des champs CMG / DPB / GC (2.4.3)
# --------------------------------------------------------------------------


def _encrypt(data, project_id, seed=None):
    seed = random.randint(0, 255) if seed is None else seed
    version = 2
    proj_key = sum(project_id.encode("ascii")) & 0xFF
    out = bytearray([seed, seed ^ version, seed ^ proj_key])
    unencrypted_byte1 = proj_key
    encrypted_byte1 = seed ^ proj_key
    encrypted_byte2 = seed ^ version
    ignored_length = (seed & 6) // 2
    stream = bytes([0] * ignored_length) + struct.pack("<I", len(data)) + data
    for b in stream:
        enc = ((encrypted_byte2 + unencrypted_byte1) & 0xFF) ^ b
        out.append(enc)
        encrypted_byte2 = encrypted_byte1
        encrypted_byte1 = enc
        unencrypted_byte1 = b
    return out.hex().upper()


# --------------------------------------------------------------------------
# Flux dir (2.3.4.2)
# --------------------------------------------------------------------------


def _rec(rec_id, payload):
    return struct.pack("<HI", rec_id, len(payload)) + payload


def _mbcs(s):
    return s.encode("cp1252")


def _utf16(s):
    return s.encode("utf-16-le")


def _dir_stream(project_name, modules):
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 3))           # SYSKIND win64
    d += _rec(0x004A, struct.pack("<I", 6))           # COMPATVERSION
    d += _rec(0x0002, struct.pack("<I", 0x409))       # LCID
    d += _rec(0x0014, struct.pack("<I", 0x409))       # LCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", CODEPAGE))    # CODEPAGE
    d += _rec(0x0004, _mbcs(project_name))            # NAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")        # DOCSTRING
    d += _rec(0x0006, b"") + _rec(0x003D, b"")        # HELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))           # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))           # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)        # VERSION
    d += _rec(0x000C, b"") + _rec(0x003C, b"")        # CONSTANTS

    refs = [
        ("stdole", "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
                   "C:\\Windows\\System32\\stdole2.tlb#OLE Automation"),
        ("Office", "*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#"
                   "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL"
                   "#Microsoft Office 16.0 Object Library"),
    ]
    for name, libid in refs:
        d += _rec(0x0016, _mbcs(name)) + _rec(0x003E, _utf16(name))
        lib = _mbcs(libid)
        d += _rec(0x000D, struct.pack("<I", len(lib)) + lib + struct.pack("<IH", 0, 0))

    d += _rec(0x000F, struct.pack("<H", len(modules)))   # MODULES count
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))         # PROJECTCOOKIE
    for mod in modules:
        n = mod["name"]
        d += _rec(0x0019, _mbcs(n))
        d += _rec(0x0047, _utf16(n))
        d += _rec(0x001A, _mbcs(n)) + _rec(0x0032, _utf16(n))
        d += _rec(0x001C, b"") + _rec(0x0048, b"")
        d += _rec(0x0031, struct.pack("<I", 0))          # MODULEOFFSET
        d += _rec(0x001E, struct.pack("<I", 0))          # HELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))     # COOKIE
        d += struct.pack("<HI", 0x0021 if mod["type"] == "std" else 0x0022, 0)
        d += struct.pack("<HI", 0x002B, 0)               # terminateur module
    d += struct.pack("<HI", 0x0010, 0)                   # terminateur dir
    return bytes(d)


# --------------------------------------------------------------------------
# Fichier composé (MS-CFB, version 3)
# --------------------------------------------------------------------------

FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


class _Entry:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data
        self.children = []
        self.left = self.right = self.child = NOSTREAM
        self.start, self.size = ENDOFCHAIN, 0
        self.sid = None


def _sort_key(e):
    return (len(e.name), e.name.upper())


def _write_cfb(root):
    entries = []

    def walk(e):
        e.sid = len(entries)
        entries.append(e)
        for c in sorted(e.children, key=_sort_key):
            walk(c)

    walk(root)
    # Arbre des frères : chaîne vers la droite, triée (arbre binaire valide).
    for e in entries:
        kids = sorted(e.children, key=_sort_key)
        if kids:
            e.child = kids[0].sid
            for a, b in zip(kids, kids[1:]):
                a.right = b.sid

    streams = [e for e in entries if e.kind == 2]
    small = [e for e in streams if len(e.data) < 4096]
    big = [e for e in streams if len(e.data) >= 4096]

    # Mini-flux
    mini = bytearray()
    minifat = []
    for e in small:
        e.size = len(e.data)
        if not e.data:
            e.start = ENDOFCHAIN
            continue
        n = (len(e.data) + 63) // 64
        e.start = len(mini) // 64
        for i in range(n):
            minifat.append(e.start + i + 1 if i < n - 1 else ENDOFCHAIN)
        mini += e.data + b"\0" * (n * 64 - len(e.data))

    sectors = []  # liste de blocs de 512 octets
    fat = []

    def alloc(data):
        if not data:
            return ENDOFCHAIN
        n = (len(data) + 511) // 512
        start = len(sectors)
        padded = data + b"\0" * (n * 512 - len(data))
        for i in range(n):
            sectors.append(padded[i * 512:(i + 1) * 512])
            fat.append(start + i + 1 if i < n - 1 else ENDOFCHAIN)
        return start

    for e in big:
        e.size = len(e.data)
        e.start = alloc(e.data)
    root.start = alloc(bytes(mini)) if mini else ENDOFCHAIN
    root.size = len(mini)
    minifat_bytes = b"".join(struct.pack("<I", x) for x in minifat)
    if minifat_bytes:
        minifat_bytes += struct.pack("<I", FREESECT) * ((-len(minifat)) % 128)
    minifat_start = alloc(minifat_bytes)
    n_minifat = (len(minifat_bytes) + 511) // 512

    dir_bytes = bytearray()
    for e in entries:
        name = _utf16(e.name) + b"\0\0"
        ent = name + b"\0" * (64 - len(name))
        ent += struct.pack("<HBB", len(name), e.kind, 1)
        ent += struct.pack("<III", e.left, e.right, e.child)
        ent += b"\0" * 16 + struct.pack("<I", 0) + b"\0" * 16
        ent += struct.pack("<IQ", e.start if e.kind != 1 else 0, e.size)
        dir_bytes += ent
    dir_bytes += (b"\0" * 64 + struct.pack("<HBB", 0, 0, 0)
                  + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM)
                  + b"\0" * 48) * ((-len(entries)) % 4)
    dir_start = alloc(bytes(dir_bytes))

    n_fat = 1
    while (len(sectors) + n_fat) > n_fat * 128:
        n_fat += 1
    fat_start = len(sectors)
    fat += [FATSECT] * n_fat
    fat += [FREESECT] * (n_fat * 128 - len(fat))
    fat_bytes = b"".join(struct.pack("<I", x) for x in fat)
    for i in range(n_fat):
        sectors.append(fat_bytes[i * 512:(i + 1) * 512])
    assert n_fat <= 109

    header = bytearray()
    header += bytes.fromhex("D0CF11E0A1B11AE1") + b"\0" * 16
    header += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6)
    header += b"\0" * 6
    header += struct.pack("<IIIIIIIII", 0, n_fat, dir_start, 0, 4096,
                          minifat_start if n_minifat else ENDOFCHAIN, n_minifat,
                          ENDOFCHAIN, 0)
    difat = [fat_start + i for i in range(n_fat)] + [FREESECT] * (109 - n_fat)
    header += b"".join(struct.pack("<I", x) for x in difat)
    assert len(header) == 512
    return bytes(header) + b"".join(sectors)


# --------------------------------------------------------------------------
# Assemblage
# --------------------------------------------------------------------------

DOC_ATTRS = {
    "workbook": "0{00020819-0000-0000-C000-000000000046}",
    "sheet": "0{00020820-0000-0000-C000-000000000046}",
}


def _module_source(mod):
    code = mod["code"].replace("\r\n", "\n").strip("\n")
    lines = [f'Attribute VB_Name = "{mod["name"]}"']
    if mod["type"] != "std":
        lines += [
            f'Attribute VB_Base = "{DOC_ATTRS[mod["type"]]}"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = True",
            "Attribute VB_TemplateDerived = False",
            "Attribute VB_Customizable = True",
        ]
    text = "\r\n".join(lines) + "\r\n" + code.replace("\n", "\r\n") + "\r\n"
    return text.encode("cp1252")


def build_vba_project(modules, path, project_name="VBAProject"):
    """modules : liste de dicts {name, type ('std'|'workbook'|'sheet'), code}."""
    project_id = "{" + str(uuid.uuid4()).upper() + "}"
    lines = [f'ID="{project_id}"']
    for m in modules:
        if m["type"] == "std":
            lines.append(f"Module={m['name']}")
        else:
            lines.append(f"Document={m['name']}/&H00000000")
    lines += [
        f'Name="{project_name}"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        f'CMG="{_encrypt(struct.pack("<I", 0), project_id)}"',
        f'DPB="{_encrypt(bytes([0]), project_id)}"',
        f'GC="{_encrypt(bytes([0xFF]), project_id)}"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    lines += [f"{m['name']}=0, 0, 0, 0, C" for m in modules]
    project_text = ("\r\n".join(lines) + "\r\n").encode("cp1252")

    wm = bytearray()
    for m in modules:
        wm += _mbcs(m["name"]) + b"\0" + _utf16(m["name"]) + b"\0\0"
    wm += b"\0\0"

    root = _Entry("Root Entry", 5)
    vba = _Entry("VBA", 1)
    root.children = [_Entry("PROJECT", 2, project_text),
                     _Entry("PROJECTwm", 2, bytes(wm)), vba]
    vba.children = [
        _Entry("_VBA_PROJECT", 2, bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x03, 0x00])),
        _Entry("dir", 2, compress(_dir_stream(project_name, modules))),
    ]
    for m in modules:
        vba.children.append(_Entry(m["name"], 2, compress(_module_source(m))))
    with open(path, "wb") as f:
        f.write(_write_cfb(root))
