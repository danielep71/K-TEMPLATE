#!/usr/bin/env python3
"""Report where an Office Open XML package was saved, by whom, and under which label.

A workbook saved in a Microsoft 365 environment can carry data from that
environment that no sheet shows. For each package given, a .xlsx, .xlsm,
.xltx, .docx or any other ZIP with [Content_Types].xml, this reports:

- path or URL: in any XML part, relationship targets of the _rels/ parts,
  external links and data connections included, a URL to a SharePoint site
  (*.sharepoint.com, or *-my.sharepoint.com for OneDrive for Business), a
  OneDrive URL, a "OneDrive - <organization>" or OneDrive folder, a file: URL,
  an extended-length path (\\\\?\\), a path on a network share (\\\\host\\share),
  and an absolute path under a user profile (C:\\Users\\<name>,
  /Users/<name>, /home/<name>); and any absolute path in x15ac:absPath, the
  folder of the last save. The VBA project (vbaProject.bin) is read for the
  same paths and URLs only, as ASCII and UTF-16LE text, both as stored and
  as its dir stream and module source decompress. A reference to a library
  under C:\\Windows or C:\\Program Files is not a finding: every VBA project
  holds one;
- person or organization: the author and last editor in docProps/core.xml,
  the company and manager in docProps/app.xml (PERSONAL_FIELDS), a userName
  attribute in any part (fileSharing in xl/workbook.xml, the headers of
  shared-workbook revisions), the people of threaded comments in
  xl/persons/, the authors of comments in xl/comments*.xml, other than the
  "tc={...}" placeholder that Excel writes for a threaded comment, and the
  users of xl/revisions/userNames.xml;
- sensitivity label: a Microsoft Purview label, as an MSIP_Label_* property
  or any other reference to one, a part under docMetadata/, such as
  docMetadata/LabelInfo.xml, and a content type or relationship that refers
  to such a part;
- SharePoint metadata: a document-library property in docProps/custom.xml,
  such as ContentTypeId, _dlc_DocId, MediaServiceImageTags or a taxonomy
  field, and any element or reference in a SharePoint metadata namespace,
  such as the property schemas that customXml/item*.xml parts hold;
- unreadable: an XML part (a .xml or .rels part) that is not well-formed,
  declares a DTD or cannot be read from the package. It cannot be checked,
  so it is a finding too.

A finding names the part, the kind and a hint, never the value it found:

    Book.xlsx: xl/workbook.xml: path or URL (x15ac:absPath holds a OneDrive for Business URL)
    Book.xlsx: docProps/custom.xml: sensitivity label (property MSIP_Label_… names a Purview sensitivity label)

Each XML part is parsed in its declared encoding, with character references
resolved, so a UTF-16 part or an escaped character hides nothing; a part
with a DTD is refused, never expanded. It reads every XML part, but no cell
content beyond paths and URLs, and it never changes a file.

tools/check_binaries.py runs it on every committed Office package, and
tools/release.py verify-published on every attached one.

Usage: python tools/check_workbook_privacy.py FILE...
Exit status: 0 when no file has a finding, 1 when one has, 2 when a file
cannot be read or is not an Office Open XML package; 2 wins over 1.
"""

from __future__ import annotations

import argparse
import io
import re
import struct
import sys
import zipfile
import zlib
from collections.abc import Iterator, Sequence
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree
from xml.parsers import expat

CONTENT_TYPES = "[Content_Types].xml"
# docProps fields that name a person or an organization, by part.
PERSONAL_FIELDS = {"docProps/core.xml": ("creator", "lastModifiedBy"),
                   "docProps/app.xml": ("Company", "Manager")}

PATH = "path or URL"
PERSON = "person or organization"
LABEL = "sensitivity label"
SHAREPOINT = "SharePoint metadata"
UNREADABLE = "unreadable"

# The path and URL forms, most specific first: the first that matches a text
# describes it.
LOCATIONS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("a OneDrive for Business URL",
     re.compile(r"[a-z0-9-]+-my\.sharepoint(?:-df)?\.(?:com|us|de|cn)\b", re.I)),
    ("a SharePoint URL",
     re.compile(r"[a-z0-9-]+\.sharepoint(?:-df)?\.(?:com|us|de|cn)\b", re.I)),
    ("a OneDrive URL", re.compile(r"\b(?:d\.docs\.live\.net|onedrive\.live\.com|1drv\.ms)\b",
                                  re.I)),
    ("a OneDrive for Business folder named after its organization",
     re.compile(r"[\\/]OneDrive(?: - |-)(?!Personal\b)\w", re.I)),
    ("a OneDrive folder", re.compile(r"[\\/]OneDrive(?:-Personal)?(?:[\\/]|$)", re.I)),
    ("a file: URL", re.compile(r"\bfile:[/\\]", re.I)),
    ("an extended-length path that starts with \\\\?\\", re.compile(r"\\\\\?\\")),
    ("an absolute path under a user profile", re.compile(
        r"\b[a-z]:[\\/]+(?:users|documents and settings)[\\/]+"
        r"(?!(?:public|default|default user|all users)(?:[\\/]|$))[^\\/:*?\"<>|\r\n]", re.I)),
    ("an absolute path under a user profile",
     re.compile(r"(?<![\w.~:/\\-])/(?:Users/(?!Shared(?:/|$))|home/)[^/\s<>\"]")),
    ("a path on a network share",
     re.compile(r"(?<![\w:/\\?.])\\\\[\w$-]+(?:\.[\w$-]+)*\\[^\\/:*?\"<>|\s]")),
)
LABEL_TEXT = re.compile(r"MSIP_Label_", re.I)
# A content type, part name or relationship of a sensitivity label part.
LABEL_PART = re.compile(r"docMetadata/|classificationlabels", re.I)
CUSTOM_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
# Document-library properties that SharePoint writes into docProps/custom.xml.
SHAREPOINT_PROPERTIES = frozenset({
    "ContentTypeId", "TaxCatchAll", "TaxCatchAllLabel", "TaxKeyword", "TaxKeywordTaxHTField",
    "lcf76f155ced4ddcb4097134ff3c332f", "ComplianceAssetId", "_ExtendedDescription",
    "SharedWithUsers", "SharedWithDetails", "TriggerFlowInfo", "TemplateUrl", "xd_ProgID",
    "xd_Signature", "_SourceUrl", "_SharedFileIndex"})
SHAREPOINT_PREFIXES = ("_dlc_", "_ip_UnifiedCompliancePolicy", "MediaService",
                       "AuthorIds_UIVersion_", "display_urn:schemas-microsoft-com:office:office#")
SHAREPOINT_NAMESPACES = ("http://schemas.microsoft.com/office/2006/metadata/",
                         "http://schemas.microsoft.com/sharepoint/",
                         "http://schemas.microsoft.com/office/infopath/2007/PartnerControls")
# The author Excel writes for the legacy copy of a threaded comment.
THREADED_AUTHOR = re.compile(r"tc=\{[0-9a-f-]+\}", re.I)
# Text in a binary part: printable ASCII, or the same as UTF-16LE.
ASCII_RUN = re.compile(rb"[\x20-\x7e]{4,}")
UTF16_RUN = re.compile(rb"(?:[\x20-\x7e]\x00){4,}")
CFB_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
FREE_SECTOR = 0xFFFFFFFF
END_OF_CHAIN = 0xFFFFFFFE


class NotAPackage(Exception):
    """The input is not a readable Office Open XML package."""


@dataclass(frozen=True)
class Finding:
    part: str
    kind: str
    hint: str

    def __str__(self) -> str:
        return f"{self.part}: {self.kind} ({self.hint})"


def local_name(tag: str) -> str:
    """An element or attribute name without its namespace."""
    return tag.rpartition("}")[2]


def namespace(tag: str) -> str:
    return tag[1:].partition("}")[0] if tag.startswith("{") else ""


class _Doctype(Exception):
    pass


def _refuse_doctype(*_: object) -> None:
    raise _Doctype


def declares_dtd(data: bytes) -> bool:
    """Whether the XML declares a document type, read in its declared encoding."""
    parser = expat.ParserCreate()
    parser.StartDoctypeDeclHandler = _refuse_doctype
    try:
        parser.Parse(data, True)
    except _Doctype:
        return True
    except expat.ExpatError:
        return False
    return False


def read_part(data: bytes) -> ElementTree.Element | None:
    """The parsed part, read in its declared encoding, or None when it is not XML.

    The parser honours a byte-order mark and the XML declaration, so a UTF-16
    part is read like a UTF-8 one, and it resolves character references. A
    part that declares a DTD is refused before any entity can be expanded.
    """
    if declares_dtd(data):
        return None
    try:
        return ElementTree.fromstring(data)
    except ElementTree.ParseError:
        return None


def describe(text: str) -> str | None:
    """The description of the first path or URL form the text holds, if any."""
    return next((label for label, pattern in LOCATIONS if pattern.search(text)), None)


def strings(element: ElementTree.Element) -> Iterator[tuple[str, str]]:
    """Where each text and attribute value of the element is, and the value."""
    tag = local_name(element.tag)
    for name, value in element.attrib.items():
        yield f"{tag}@{local_name(name)}", value
    for text in (element.text, element.tail):
        if text and text.strip():
            yield tag, text


def path_hints(root: ElementTree.Element) -> Iterator[tuple[str, str]]:
    for element in root.iter():
        if local_name(element.tag) == "absPath":
            if (element.get("url") or "").strip():
                found = describe(element.get("url") or "") or "an absolute path"
                yield PATH, f"x15ac:absPath holds {found}"
            continue
        for where, value in strings(element):
            described = describe(value)
            if described:
                yield PATH, f"{where} holds {described}"


def is_comments(part: str) -> bool:
    return re.fullmatch(r"xl/comments[^/]*\.xml", part, re.I) is not None


def names_a_person(part: str, tag: str, attribute: str) -> bool:
    """Whether this attribute of this element names a person, in this part."""
    lower = part.lower()
    if attribute == "userName":
        return True
    if lower.startswith("xl/persons/"):
        return tag == "person" and attribute in ("displayName", "userId")
    return lower.startswith("xl/revisions/") and tag == "userInfo" and attribute == "name"


def named_people(part: str, element: ElementTree.Element) -> Iterator[str]:
    """Where the element names a person, given the part it is in."""
    tag = local_name(element.tag)
    for name, value in element.attrib.items():
        if value.strip() and names_a_person(part, tag, local_name(name)):
            yield f"{tag}@{local_name(name)}"
    text = (element.text or "").strip()
    if text and (tag in PERSONAL_FIELDS.get(part, ())
                 or is_comments(part) and tag == "author" and not THREADED_AUTHOR.fullmatch(text)):
        yield tag


def person_hints(part: str, root: ElementTree.Element) -> Iterator[tuple[str, str]]:
    for element in root.iter():
        for where in named_people(part, element):
            yield PERSON, f"{where} is not blank"


def custom_property_hint(name: str) -> tuple[str, str] | None:
    """The finding for a custom document property of this name, if any."""
    if LABEL_TEXT.match(name):
        return LABEL, "property MSIP_Label_… names a Purview sensitivity label"
    if name in SHAREPOINT_PROPERTIES:
        return SHAREPOINT, f"property {name} is a SharePoint library property"
    prefix = next((prefix for prefix in SHAREPOINT_PREFIXES if name.startswith(prefix)), None)
    if prefix is not None:
        return SHAREPOINT, f"property {prefix}… is a SharePoint library property"
    return None


def metadata_hints(root: ElementTree.Element) -> Iterator[tuple[str, str]]:
    """Purview labels and SharePoint metadata in the part."""
    for element in root.iter():
        tag = local_name(element.tag)
        if tag == "property" and namespace(element.tag) == CUSTOM_PROPERTIES:
            found = custom_property_hint(element.get("name") or "")
            if found:
                yield found
                continue
        if namespace(element.tag).startswith(SHAREPOINT_NAMESPACES):
            yield SHAREPOINT, "holds elements of a SharePoint metadata namespace"
        for where, value in strings(element):
            if LABEL_TEXT.search(value):
                yield LABEL, f"{where} refers to a Purview sensitivity label (MSIP_Label_…)"
            elif LABEL_PART.search(value):
                yield LABEL, f"{where} refers to a Purview sensitivity label part"
            elif value.strip().startswith(SHAREPOINT_NAMESPACES):
                yield SHAREPOINT, f"{where} refers to a SharePoint metadata schema"


def xml_hints(part: str, data: bytes) -> Iterator[tuple[str, str]]:
    root = read_part(data)
    if root is None:
        reason = "declares a DTD" if declares_dtd(data) else "is not well-formed XML"
        yield UNREADABLE, f"{reason}, so it cannot be checked"
        return
    yield from path_hints(root)
    yield from person_hints(part, root)
    yield from metadata_hints(root)


# ---------------------------------------------------------------------------
# The VBA project: a compound file [MS-CFB] whose dir stream and module source
# are compressed [MS-OVBA]. Anything that cannot be read is left to the scan
# of the bytes as stored.

def decompress(data: bytes) -> bytes:
    """[MS-OVBA] 2.4.1.3.1: decompress a CompressedContainer; ValueError if malformed."""
    if not data or data[0] != 0x01:
        raise ValueError("no compressed container")
    out, pos = bytearray(), 1
    while pos + 2 <= len(data):
        header = int.from_bytes(data[pos:pos + 2], "little")
        end, pos = min(pos + (header & 0x0FFF) + 3, len(data)), pos + 2
        if header >> 12 & 0x7 != 0b011:
            raise ValueError("bad chunk signature")
        if not header & 0x8000:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        start = len(out)
        while pos < end:
            flags, pos = data[pos], pos + 1
            for bit in range(8):
                if pos >= end:
                    break
                if not flags >> bit & 1:
                    out.append(data[pos])
                    pos += 1
                    continue
                token, pos = int.from_bytes(data[pos:pos + 2], "little"), pos + 2
                bits = max((len(out) - start - 1).bit_length(), 4)
                source = len(out) - ((token >> (16 - bits)) + 1)
                if source < start:
                    raise ValueError("copy token before the chunk")
                for index in range((token & (0xFFFF >> bits)) + 3):
                    out.append(out[source + index])
    return bytes(out)


class CompoundFile:
    """The streams of an [MS-CFB] compound file, by name in lower case."""

    def __init__(self, blob: bytes) -> None:
        if blob[:8] != CFB_SIGNATURE:
            raise ValueError("not a compound file")
        self.blob = blob
        self.size = 1 << struct.unpack_from("<H", blob, 30)[0]
        self.mini_size = 1 << struct.unpack_from("<H", blob, 32)[0]
        directory_start, = struct.unpack_from("<I", blob, 48)
        self.cutoff, minifat_start, _, difat_start, difat_count = \
            struct.unpack_from("<IIIII", blob, 56)
        fat_sectors = [sector for sector in struct.unpack_from("<109I", blob, 76)
                       if sector != FREE_SECTOR]
        for sector in self.chain_of_difat(difat_start, difat_count):
            entries = struct.unpack(f"<{self.size // 4}I", self.sector(sector))
            fat_sectors += [entry for entry in entries[:-1] if entry != FREE_SECTOR]
        self.fat = self.table(b"".join(self.sector(sector) for sector in fat_sectors))
        directory = self.read(directory_start, self.fat)
        self.entries = [directory[index:index + 128] for index in range(0, len(directory), 128)]
        root_start, root_size = self.position(self.entries[0])
        self.mini_stream = self.read(root_start, self.fat)[:root_size]
        self.minifat = self.table(self.read(minifat_start, self.fat))

    @staticmethod
    def table(data: bytes) -> list[int]:
        return list(struct.unpack(f"<{len(data) // 4}I", data[:len(data) // 4 * 4]))

    def chain_of_difat(self, start: int, count: int) -> Iterator[int]:
        for _ in range(count):
            if start in (END_OF_CHAIN, FREE_SECTOR):
                return
            yield start
            start, = struct.unpack_from("<I", self.sector(start), self.size - 4)

    def sector(self, index: int) -> bytes:
        offset = (index + 1) * self.size
        if offset >= len(self.blob):
            raise ValueError("sector past the end of the file")
        return self.blob[offset:offset + self.size]

    @staticmethod
    def chain(start: int, table: list[int]) -> list[int]:
        found: list[int] = []
        while start not in (END_OF_CHAIN, FREE_SECTOR):
            if start >= len(table) or len(found) > len(table):
                raise ValueError("broken sector chain")
            found.append(start)
            start = table[start]
        return found

    def read(self, start: int, table: list[int]) -> bytes:
        return b"".join(self.sector(sector) for sector in self.chain(start, table))

    @staticmethod
    def position(entry: bytes) -> tuple[int, int]:
        """The first sector and the size of a directory entry's stream."""
        start, size = struct.unpack_from("<IQ", entry, 116)
        return start, size & 0xFFFFFFFF

    def streams(self) -> dict[str, list[bytes]]:
        found: dict[str, list[bytes]] = {}
        for entry in self.entries:
            if len(entry) < 128 or entry[66] != 2:  # 2: a stream
                continue
            length = min(struct.unpack_from("<H", entry, 64)[0], 64)
            name = entry[:max(length - 2, 0)].decode("utf-16-le", "replace").lower()
            start, size = self.position(entry)
            if size < self.cutoff:
                data = b"".join(self.mini_stream[sector * self.mini_size:
                                                 (sector + 1) * self.mini_size]
                                for sector in self.chain(start, self.minifat))
            else:
                data = self.read(start, self.fat)
            found.setdefault(name, []).append(data[:size])
        return found


def module_offsets(directory: bytes) -> dict[str, int]:
    """The stream name and source offset of each module, from the decompressed dir stream."""
    offsets: dict[str, int] = {}
    pos, name = 0, ""
    while pos + 6 <= len(directory):
        record, size = struct.unpack_from("<HI", directory, pos)
        payload = directory[pos + 6:pos + 6 + size]
        # PROJECTVERSION: its size field is a reserved 4, then 6 bytes follow.
        pos += 6 + (6 if record == 0x0009 else size)
        if record == 0x001A:
            name = payload.decode("latin-1").lower()
        elif record == 0x0031 and len(payload) == 4:
            offsets[name] = int.from_bytes(payload, "little")
    return offsets


def compressed_texts(blob: bytes) -> Iterator[bytes]:
    """The dir stream and the source of each module, decompressed, as far as they read."""
    try:
        streams = CompoundFile(blob).streams()
        directory = decompress(streams["dir"][0])
    except (ValueError, KeyError, IndexError, struct.error):
        return
    yield directory
    for name, offset in module_offsets(directory).items():
        for data in streams.get(name, []):
            try:
                yield decompress(data[offset:])
            except (ValueError, IndexError):
                continue


def binary_hints(data: bytes) -> Iterator[tuple[str, str]]:
    """Paths and URLs in the text of a VBA project, as stored and decompressed."""
    for blob in (data, *compressed_texts(data)):
        for pattern, encoding in ((ASCII_RUN, "ascii"), (UTF16_RUN, "utf-16-le")):
            for run in pattern.finditer(blob):
                found = describe(run.group().decode(encoding))
                if found:
                    yield PATH, f"VBA project text holds {found}"


# ---------------------------------------------------------------------------

def read_package(data: bytes) -> dict[str, bytes | None]:
    """Each part of the package by name, None for one that cannot be read.

    NotAPackage when the data is not a ZIP with [Content_Types].xml.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as package:
            names = [info.filename for info in package.infolist() if not info.is_dir()]
            if CONTENT_TYPES not in names:
                raise NotAPackage(f"a ZIP without {CONTENT_TYPES}, not an Office Open XML "
                                  "package")
            parts: dict[str, bytes | None] = {}
            for name in names:
                try:
                    parts[name] = package.read(name)
                except (zipfile.BadZipFile, OSError, EOFError, RuntimeError, ValueError,
                        NotImplementedError, zlib.error):
                    parts[name] = None
            return parts
    except (zipfile.BadZipFile, OSError, EOFError, ValueError) as error:
        raise NotAPackage(f"not a ZIP, so not an Office Open XML package ({error})") from None


def is_package(data: bytes) -> bool:
    try:
        read_package(data)
    except NotAPackage:
        return False
    return True


def part_hints(name: str, data: bytes | None) -> Iterator[tuple[str, str]]:
    lower = name.lower()
    if lower.startswith("docmetadata/"):
        yield LABEL, "a Purview sensitivity label part"
    if data is None:
        yield UNREADABLE, "cannot be read from the package, so it cannot be checked"
    elif lower.endswith((".xml", ".rels")):
        yield from xml_hints(name, data)
    elif lower.endswith("vbaproject.bin"):
        yield from binary_hints(data)


def scan(data: bytes) -> list[Finding]:
    """The findings in the package, each once, in part order; NotAPackage if it is none."""
    return list(dict.fromkeys(Finding(name, kind, hint)
                              for name, content in read_package(data).items()
                              for kind, hint in part_hints(name, content)))


def scan_file(path: Path) -> list[Finding]:
    """scan() of a file; OSError or NotAPackage when it cannot be checked."""
    return scan(path.read_bytes())


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("files", nargs="+", type=Path, metavar="FILE",
                        help="an Office Open XML package, such as a .xlsx or .xlsm workbook")
    args = parser.parse_args(argv)
    found, unreadable = 0, 0
    for path in args.files:
        try:
            findings = scan_file(path)
        except (OSError, NotAPackage) as error:
            unreadable += 1
            print(f"{path}: cannot be checked: {error}")
            continue
        found += len(findings)
        for finding in findings:
            print(f"{path}: {finding}")
    status = "ERROR" if unreadable else "FAIL" if found else "PASS"
    print(f"workbook-privacy: {status} ({found} finding(s) in {len(args.files)} file(s), "
          f"{unreadable} unreadable)")
    return 2 if unreadable else 1 if found else 0


if __name__ == "__main__":
    sys.exit(main())
