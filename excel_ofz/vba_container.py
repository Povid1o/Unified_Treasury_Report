"""Сборка `vbaProject.bin` — проекта VBA — без Excel и без Windows.

Нужно, чтобы готовый .xlsm собирался на macOS одной командой. Реализованы две
спецификации Microsoft:

* **[MS-OVBA] 2.4.1** — алгоритм сжатия потоков `dir` и модулей
  (`compress_container`).
* **[MS-CFB]** — контейнер OLE Compound File, в котором лежит проект
  (`build_cfb`).

Модуль ничего не знает про отчёт ОФЗ: на вход — имя проекта и список модулей,
на выходе — байты `vbaProject.bin`.
"""
from __future__ import annotations

import struct
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

# ── [MS-OVBA] 2.4.1: CompressedContainer ─────────────────────────────────────

SIGNATURE_BYTE = 0x01
MAX_CHUNK_DATA = 0x0FFF  # 12 бит под размер -> данные чанка не больше 4095 байт
CHUNK_DECOMPRESSED_MAX = 4096


def _ceiling_log2(value: int) -> int:
    """CeilingLog2 из [MS-OVBA] 2.4.1.3.19.1: минимальное n, при котором 2^n >= value."""
    return (value - 1).bit_length() if value > 1 else 0


def _copy_token_split(difference: int) -> Tuple[int, int]:
    """CopyTokenHelp: сколько бит занимает смещение и какова максимальная длина."""
    bit_count = max(4, _ceiling_log2(difference))
    length_mask = 0xFFFF >> bit_count
    return bit_count, length_mask


def _find_match(data: bytes, window_start: int, pos: int, limit: int,
                index: Dict[bytes, List[int]]) -> Tuple[int, int]:
    """Самое длинное совпадение для data[pos:] в пределах окна [window_start, pos)."""
    best_len, best_off = 0, 0
    if pos + 3 > limit:
        return 0, 0

    checked = 0
    for cand in reversed(index.get(data[pos:pos + 3], ())):
        if cand >= pos:
            continue                      # позиция ещё не «разжата», ссылаться нельзя
        if cand < window_start:
            break                         # вышли за начало чанка — окно кончилось
        checked += 1
        if checked > 64:
            break                         # дальше выигрыш копеечный, а перебор дорогой

        length = 0
        # Совпадение может заходить за текущую позицию: декомпрессор копирует
        # побайтно, поэтому так кодируются повторяющиеся последовательности.
        while pos + length < limit and data[cand + length] == data[pos + length]:
            length += 1
        if length > best_len:
            best_len, best_off = length, pos - cand
    return best_off, best_len


def _compress_chunk(data: bytes, start: int, hard_end: int) -> Tuple[bytes, int]:
    """Кодирует один чанк. Возвращает (байты чанка без заголовка, сколько байт съедено)."""
    out = bytearray()
    end = min(start + CHUNK_DECOMPRESSED_MAX, hard_end)
    pos = start

    index: Dict[bytes, List[int]] = {}
    for i in range(start, min(end, hard_end) - 2):
        index.setdefault(data[i:i + 3], []).append(i)

    while pos < end:
        seq = bytearray()
        flags = 0
        cur = pos

        for bit in range(8):
            if cur >= end:
                break
            offset, length = _find_match(data, start, cur, end, index)
            if length >= 3:
                bit_count, length_mask = _copy_token_split(cur - start)
                length = min(length, length_mask + 3)
                token = ((offset - 1) << (16 - bit_count)) | (length - 3)
                seq += struct.pack("<H", token)
                flags |= 1 << bit
                cur += length
            else:
                seq.append(data[cur])
                cur += 1

        # Чанк не может занять больше 4095 байт — не влезло, значит он закончился.
        if len(out) + 1 + len(seq) > MAX_CHUNK_DATA:
            break
        out.append(flags)
        out += seq
        pos = cur

    return bytes(out), pos - start


def compress_container(data: bytes) -> bytes:
    """Сжимает поток в CompressedContainer ([MS-OVBA] 2.4.1.1)."""
    out = bytearray([SIGNATURE_BYTE])
    pos = 0
    while pos < len(data):
        chunk, consumed = _compress_chunk(data, pos, len(data))
        if consumed == 0:
            raise ValueError("Сжатие не продвинулось — повреждённые входные данные")
        # Бит 15 = 1 (чанк сжат), биты 12-14 = 0b011, биты 0-11 = размер чанка - 3.
        header = 0xB000 | ((len(chunk) + 2 - 3) & MAX_CHUNK_DATA)
        out += struct.pack("<H", header) + chunk
        pos += consumed
    return bytes(out)


def decompress_container(data: bytes) -> bytes:
    """Обратная операция — только для самопроверки сжатия."""
    if not data or data[0] != SIGNATURE_BYTE:
        raise ValueError("Неверная сигнатура CompressedContainer")
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        pos += 2
        size = (header & MAX_CHUNK_DATA) + 3
        compressed = bool(header & 0x8000)
        chunk_end = pos + size - 2

        if not compressed:
            out += data[pos:chunk_end]
            pos = chunk_end
            continue

        chunk_start = len(out)
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if flags & (1 << bit):
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    bit_count, length_mask = _copy_token_split(len(out) - chunk_start)
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    src = len(out) - offset
                    for i in range(length):
                        out.append(out[src + i])
                else:
                    out.append(data[pos])
                    pos += 1
    return bytes(out)


# ── [MS-CFB]: контейнер OLE ──────────────────────────────────────────────────

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_CUTOFF = 4096

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

TYPE_STORAGE = 1
TYPE_STREAM = 2
TYPE_ROOT = 5


@dataclass
class _Entry:
    name: str
    kind: int
    data: bytes = b""
    children: List[str] = field(default_factory=list)
    left: int = NOSTREAM
    right: int = NOSTREAM
    child: int = NOSTREAM
    start: int = ENDOFCHAIN
    size: int = 0


def _dir_sort_key(name: str) -> Tuple[int, str]:
    """Порядок записей в дереве CFB: сначала по длине имени, потом без учёта регистра."""
    return len(name), name.upper()


def _build_tree(ids: List[int], entries: List[_Entry]) -> int:
    """Строит сбалансированное дерево из отсортированного списка; корень возвращается."""
    if not ids:
        return NOSTREAM
    mid = len(ids) // 2
    node = ids[mid]
    entries[node].left = _build_tree(ids[:mid], entries)
    entries[node].right = _build_tree(ids[mid + 1:], entries)
    return node


def _chain(fat: List[int], sectors: List[bytes], data: bytes) -> int:
    """Кладёт данные в цепочку секторов, возвращает номер первого сектора."""
    if not data:
        return ENDOFCHAIN
    first = len(sectors)
    blocks = [data[i:i + SECTOR_SIZE] for i in range(0, len(data), SECTOR_SIZE)]
    for i, block in enumerate(blocks):
        sectors.append(block.ljust(SECTOR_SIZE, b"\x00"))
        fat.append(first + i + 1 if i + 1 < len(blocks) else ENDOFCHAIN)
    return first


def build_cfb(tree: Dict[str, object]) -> bytes:
    """Собирает CFB-файл.

    `tree` — вложенные словари: значение-словарь становится хранилищем,
    значение-bytes — потоком. Порядок ключей не важен.
    """
    entries: List[_Entry] = [_Entry("Root Entry", TYPE_ROOT)]

    def add(node: Dict[str, object], parent: int) -> None:
        child_ids: List[int] = []
        for name in sorted(node.keys(), key=_dir_sort_key):
            value = node[name]
            if isinstance(value, dict):
                entries.append(_Entry(name, TYPE_STORAGE))
                idx = len(entries) - 1
                add(value, idx)
            else:
                entries.append(_Entry(name, TYPE_STREAM, data=value))
                idx = len(entries) - 1
            child_ids.append(idx)
        entries[parent].child = _build_tree(child_ids, entries)

    add(tree, 0)

    # 1. Мелкие потоки уходят в mini-stream, крупные — в обычные секторы.
    fat: List[int] = []
    sectors: List[bytes] = []
    mini_fat: List[int] = []
    mini_stream = bytearray()

    for entry in entries:
        if entry.kind != TYPE_STREAM:
            continue
        entry.size = len(entry.data)
        if not entry.data:
            entry.start = ENDOFCHAIN
        elif entry.size < MINI_CUTOFF:
            first = len(mini_fat)
            blocks = [entry.data[i:i + MINI_SECTOR_SIZE]
                      for i in range(0, entry.size, MINI_SECTOR_SIZE)]
            for i, block in enumerate(blocks):
                mini_stream += block.ljust(MINI_SECTOR_SIZE, b"\x00")
                mini_fat.append(first + i + 1 if i + 1 < len(blocks) else ENDOFCHAIN)
            entry.start = first
        else:
            entry.start = _chain(fat, sectors, entry.data)

    # 2. Сам mini-stream хранится как обычный поток корневой записи.
    entries[0].start = _chain(fat, sectors, bytes(mini_stream))
    entries[0].size = len(mini_stream)

    # 3. MiniFAT.
    mini_fat_bytes = b"".join(struct.pack("<I", v) for v in mini_fat)
    if mini_fat_bytes:
        pad = (-len(mini_fat_bytes)) % SECTOR_SIZE
        mini_fat_bytes += struct.pack("<I", FREESECT) * (pad // 4)
    mini_fat_start = _chain(fat, sectors, mini_fat_bytes) if mini_fat_bytes else ENDOFCHAIN
    mini_fat_count = len(mini_fat_bytes) // SECTOR_SIZE

    # 4. Каталог.
    dir_bytes = b"".join(_dir_entry_bytes(e) for e in entries)
    pad = (-len(dir_bytes)) % SECTOR_SIZE
    if pad:
        dir_bytes += _empty_dir_entry() * (pad // 128)
    dir_start = _chain(fat, sectors, dir_bytes)

    # 5. FAT: сами секторы FAT тоже занимают места в FAT — подбираем количество.
    fat_sector_count = 1
    while True:
        total = len(sectors) + fat_sector_count
        needed = -(-total * 4 // SECTOR_SIZE)
        if needed <= fat_sector_count:
            break
        fat_sector_count = needed
    if fat_sector_count > 109:
        raise ValueError("Файл слишком велик: потребовался бы DIFAT-сектор")

    fat_positions = list(range(len(sectors), len(sectors) + fat_sector_count))
    full_fat = fat + [FATSECT] * fat_sector_count
    entries_per_sector = SECTOR_SIZE // 4
    full_fat += [FREESECT] * (fat_sector_count * entries_per_sector - len(full_fat))
    fat_bytes = b"".join(struct.pack("<I", v) for v in full_fat)
    for i in range(fat_sector_count):
        sectors.append(fat_bytes[i * SECTOR_SIZE:(i + 1) * SECTOR_SIZE])

    # 6. Заголовок.
    difat = fat_positions + [FREESECT] * (109 - len(fat_positions))
    header = bytearray()
    header += b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"      # сигнатура
    header += b"\x00" * 16                              # CLSID
    header += struct.pack("<HH", 0x003E, 0x0003)        # версия 3
    header += struct.pack("<H", 0xFFFE)                 # порядок байт
    header += struct.pack("<HH", 9, 6)                  # сектор 512, мини-сектор 64
    header += b"\x00" * 6
    header += struct.pack("<I", 0)                      # число секторов каталога (v3: 0)
    header += struct.pack("<I", fat_sector_count)
    header += struct.pack("<I", dir_start)
    header += struct.pack("<I", 0)                      # сигнатура транзакции
    header += struct.pack("<I", MINI_CUTOFF)
    header += struct.pack("<I", mini_fat_start)
    header += struct.pack("<I", mini_fat_count)
    header += struct.pack("<I", ENDOFCHAIN)             # первый DIFAT-сектор
    header += struct.pack("<I", 0)                      # число DIFAT-секторов
    header += b"".join(struct.pack("<I", v) for v in difat)
    assert len(header) == SECTOR_SIZE, len(header)

    return bytes(header) + b"".join(sectors)


def _dir_entry_bytes(entry: _Entry) -> bytes:
    name = entry.name.encode("utf-16-le") + b"\x00\x00"
    if len(name) > 64:
        raise ValueError(f"Слишком длинное имя записи CFB: {entry.name}")
    out = bytearray(128)
    out[0:len(name)] = name
    struct.pack_into("<H", out, 0x40, len(name))
    out[0x42] = entry.kind
    out[0x43] = 1                                        # чёрный узел
    struct.pack_into("<III", out, 0x44, entry.left, entry.right, entry.child)
    struct.pack_into("<I", out, 0x74, entry.start)
    struct.pack_into("<Q", out, 0x78, entry.size)
    return bytes(out)


def _empty_dir_entry() -> bytes:
    out = bytearray(128)
    struct.pack_into("<III", out, 0x44, NOSTREAM, NOSTREAM, NOSTREAM)
    return bytes(out)


# ── [MS-OVBA] 2.3.4.2: поток `dir` ───────────────────────────────────────────

MODULE_STANDARD = 0x0021
MODULE_DOCUMENT = 0x0022


# GUID типа, который подставляется в атрибут VB_Base модуля документа.
DOC_BASE_WORKBOOK = "{00020819-0000-0000-C000-000000000046}"
DOC_BASE_WORKSHEET = "{00020820-0000-0000-C000-000000000046}"


@dataclass
class VbaModule:
    name: str
    source: str
    kind: int = MODULE_STANDARD
    doc_base: str = DOC_BASE_WORKSHEET


def _rec(rec_id: int, payload: bytes = b"") -> bytes:
    return struct.pack("<HI", rec_id, len(payload)) + payload


def _mbcs(text: str, code_page: int) -> bytes:
    return text.encode(f"cp{code_page}")


def _reference_registered(name: str, libid: str, code_page: int) -> bytes:
    """REFERENCEREGISTERED: ссылка на зарегистрированную библиотеку типов."""
    out = _rec(0x0016, _mbcs(name, code_page))
    out += _rec(0x003E, name.encode("utf-16-le"))
    lib = libid.encode("ascii")
    body = struct.pack("<I", len(lib)) + lib + struct.pack("<IH", 0, 0)
    out += struct.pack("<HI", 0x000D, len(body)) + body
    return out


REFERENCES = [
    ("stdole",
     r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\SysWOW64\stdole2.tlb#OLE Automation"),
    ("Office",
     r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.8#0"
     r"#C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL"
     r"#Microsoft Office 16.0 Object Library"),
]


def build_dir_stream(project_name: str, modules: List[VbaModule], code_page: int = 1251) -> bytes:
    out = bytearray()

    # PROJECTINFORMATION
    out += _rec(0x0001, struct.pack("<I", 0x00000001))          # SysKind: Win32
    out += _rec(0x0002, struct.pack("<I", 0x00000419))          # Lcid: ru-RU
    out += _rec(0x0014, struct.pack("<I", 0x00000419))          # LcidInvoke
    out += _rec(0x0003, struct.pack("<H", code_page))
    out += _rec(0x0004, _mbcs(project_name, code_page))
    out += _rec(0x0005, b"") + _rec(0x0040, b"")                # DocString + Unicode
    out += _rec(0x0006, b"") + _rec(0x003D, b"")                # HelpFile1 + HelpFile2
    out += _rec(0x0007, struct.pack("<I", 0))                   # HelpContext
    out += _rec(0x0008, struct.pack("<I", 0))                   # LibFlags
    # VersionRecord: поле размера фиксировано (4), а следом идут 6 байт версии.
    out += struct.pack("<HI", 0x0009, 4) + struct.pack("<IH", 1, 0)
    out += _rec(0x000C, b"") + _rec(0x003C, b"")                # Constants + Unicode

    # PROJECTREFERENCES
    for name, libid in REFERENCES:
        out += _reference_registered(name, libid, code_page)

    # PROJECTMODULES
    out += _rec(0x000F, struct.pack("<H", len(modules)))
    out += _rec(0x0013, struct.pack("<H", 0xFFFF))              # ProjectCookie

    for module in modules:
        out += _rec(0x0019, _mbcs(module.name, code_page))      # ModuleName
        out += _rec(0x0047, module.name.encode("utf-16-le"))    # ModuleNameUnicode
        out += _rec(0x001A, _mbcs(module.name, code_page))      # ModuleStreamName
        out += _rec(0x0032, module.name.encode("utf-16-le"))
        out += _rec(0x001C, b"") + _rec(0x0048, b"")            # ModuleDocString
        # ModuleOffset: сколько байт кэша лежит перед сжатым исходником.
        # Кэш не пишем, поэтому 0 — поток целиком является контейнером.
        out += _rec(0x0031, struct.pack("<I", 0))
        out += _rec(0x001E, struct.pack("<I", 0))               # ModuleHelpContext
        out += _rec(0x002C, struct.pack("<H", 0xFFFF))          # ModuleCookie
        # MODULETYPE: и 0x0021 (обычный модуль), и 0x0022 (модуль документа)
        # идут с нулевым размером — признак несёт сам идентификатор записи.
        out += _rec(module.kind, b"")
        out += _rec(0x002B, b"")                                # Terminator

    out += struct.pack("<HI", 0x0010, 0)                        # Terminator
    return bytes(out)


def build_project_stream(project_name: str, modules: List[VbaModule], code_page: int = 1251) -> bytes:
    """Текстовый поток PROJECT. CMG/DPB/GC не пишем — проект без пароля."""
    lines = [f'ID="{{{_stable_guid(project_name)}}}"']
    for module in modules:
        if module.kind == MODULE_DOCUMENT:
            lines.append(f"Document={module.name}/&H00000000")
        else:
            lines.append(f"Module={module.name}")
    lines += [
        f'Name="{project_name}"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for module in modules:
        lines.append(f"{module.name}=0, 0, 0, 0, C")
    return ("\r\n".join(lines) + "\r\n").encode(f"cp{code_page}")


def build_projectwm_stream(modules: List[VbaModule], code_page: int = 1251) -> bytes:
    """PROJECTwm: пары «имя модуля в MBCS» -> «имя в UTF-16»."""
    out = bytearray()
    for module in modules:
        out += _mbcs(module.name, code_page) + b"\x00"
        out += module.name.encode("utf-16-le") + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


def _stable_guid(seed: str) -> str:
    """Детерминированный GUID проекта — чтобы пересборка не меняла файл без нужды."""
    import hashlib

    h = hashlib.md5(seed.encode("utf-8")).hexdigest().upper()
    return f"{h[0:8]}-{h[8:12]}-{h[12:16]}-{h[16:20]}-{h[20:32]}"


def build_vba_project(project_name: str, modules: List[VbaModule],
                      code_page: int = 1251) -> bytes:
    """Собирает готовый `vbaProject.bin`."""
    vba_storage: Dict[str, object] = {
        "_VBA_PROJECT": b"\xCC\x61\xFF\xFF\x00\x00\x00",
        "dir": compress_container(build_dir_stream(project_name, modules, code_page)),
    }
    for module in modules:
        source = _module_source_bytes(module, code_page)
        vba_storage[module.name] = compress_container(source)

    return build_cfb({
        "VBA": vba_storage,
        "PROJECT": build_project_stream(project_name, modules, code_page),
        "PROJECTwm": build_projectwm_stream(modules, code_page),
    })


def _module_source_bytes(module: VbaModule, code_page: int) -> bytes:
    """Исходник модуля с обязательными строками Attribute в начале."""
    text = module.source.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    # Attribute VB_Name из .bas-файла заменяем на свой набор атрибутов.
    while lines and lines[0].startswith("Attribute "):
        lines.pop(0)

    header = [f'Attribute VB_Name = "{module.name}"']
    if module.kind == MODULE_DOCUMENT:
        header += [
            f'Attribute VB_Base = "0{module.doc_base}"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = True",
            "Attribute VB_TemplateDerived = False",
            "Attribute VB_Customizable = True",
        ]
    body = "\r\n".join(header + lines)
    return body.encode(f"cp{code_page}", errors="replace")
