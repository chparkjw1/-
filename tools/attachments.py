"""첨부파일 내용을 읽을 수 있는 텍스트나 그림으로 바꾼다.

지원 형식
    한글(.hwp 5.x, .hwpx), 워드(.docx), 엑셀(.xlsx, .xls*), 파워포인트(.pptx),
    PDF(텍스트), 텍스트(.txt, .csv), 압축(.zip 목록), 메일(.eml), 그림(.jpg .png 등)
    * .xls 는 xlrd 가 설치된 경우에만 읽는다.
"""
import email
import io
import re
import struct
import zipfile
import zlib
from email import policy
from xml.etree import ElementTree as ET

from mailcommon import clean_text, get_body, html_to_text

IMAGE_EXT = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp"}
TEXT_EXT = {".txt", ".csv", ".log", ".md"}
XLSX_MAX_ROWS = 400
XLSX_MAX_COLS = 30


def kind_of(name, ctype, data):
    ext = ("." + name.rsplit(".", 1)[-1].lower()) if "." in name else ""
    head = data[:8]
    if ext in IMAGE_EXT or (ctype.startswith("image/") and not ext):
        return "image"
    if ext == ".pdf" or ctype == "application/pdf" or head.startswith(b"%PDF"):
        return "pdf"
    if ext in (".hwp", ".hwpx"):
        if head.startswith(b"PK\x03\x04"):
            return "hwpx"
        if head.startswith(b"\xd0\xcf\x11\xe0"):
            return "hwp"
        return "hwp-old"
    if ext == ".docx":
        return "docx"
    if ext in (".xlsx", ".xlsm"):
        return "xlsx"
    if ext == ".xls":
        return "xls"
    if ext == ".pptx":
        return "pptx"
    if ext in TEXT_EXT or ctype.startswith("text/plain"):
        return "text"
    if ext in (".htm", ".html") or ctype == "text/html":
        return "html"
    if ext == ".zip":
        return "zip"
    if ext == ".eml" or ctype == "message/rfc822":
        return "eml"
    return "other"


KIND_LABEL = {
    "image": "그림", "pdf": "PDF", "hwp": "한글 문서", "hwpx": "한글 문서(hwpx)",
    "hwp-old": "한글 문서(구버전)", "docx": "워드 문서", "xlsx": "엑셀", "xls": "엑셀(97-2003)",
    "pptx": "파워포인트", "text": "텍스트", "html": "웹 문서", "zip": "압축파일",
    "eml": "메일", "other": "기타",
}


# ---------------------------------------------------------------- 한글 (hwp 5.x)
HWPTAG_PARA_TEXT = 67
# 8칸(16바이트)을 차지하는 한글 문서의 조판 부호. 나머지 32 미만 부호는 1칸이다.
HWP_WIDE_CTRL = {1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23}


def _hwp_para(b):
    n = len(b) // 2
    codes = struct.unpack_from(f"<{n}H", b)
    keep = []
    j = 0
    while j < n:
        c = codes[j]
        if c < 32:
            if c in HWP_WIDE_CTRL:
                if c == 9:
                    keep.append(0x09)
                j += 8
                continue
            if c == 10:
                keep.append(0x0A)
            elif c in (30, 31):
                keep.append(0x20)
            j += 1
            continue
        keep.append(c)
        j += 1
    return struct.pack(f"<{len(keep)}H", *keep).decode("utf-16-le", "replace")


def hwp_text(data):
    import olefile

    ole = olefile.OleFileIO(io.BytesIO(data))
    try:
        flags = struct.unpack_from("<I", ole.openstream("FileHeader").read(), 36)[0]
        if flags & 2:
            raise ValueError("암호가 걸린 한글 문서라 내용을 읽을 수 없습니다")
        if flags & 4:
            raise ValueError("배포용(읽기 전용) 한글 문서라 내용을 읽을 수 없습니다")
        compressed = bool(flags & 1)
        sections = sorted(
            (e for e in ole.listdir() if len(e) == 2 and e[0] == "BodyText" and e[1].startswith("Section")),
            key=lambda e: int(e[1][7:] or 0),
        )
        paras = []
        for e in sections:
            raw = ole.openstream(e).read()
            if compressed:
                raw = zlib.decompress(raw, -15)
            i, size = 0, len(raw)
            while i + 4 <= size:
                h = struct.unpack_from("<I", raw, i)[0]
                i += 4
                tag, length = h & 0x3FF, (h >> 20) & 0xFFF
                if length == 0xFFF:
                    length = struct.unpack_from("<I", raw, i)[0]
                    i += 4
                if tag == HWPTAG_PARA_TEXT:
                    paras.append(_hwp_para(raw[i:i + length]))
                i += length
        return "\n".join(p.rstrip() for p in paras)
    finally:
        ole.close()


# ---------------------------------------------------------------- XML 기반 문서
def _local(tag):
    return tag.rsplit("}", 1)[-1]


def xml_text(xml_bytes, para="p", texts=("t",), tabs=("tab",), breaks=("br", "lineBreak", "cr")):
    root = ET.fromstring(xml_bytes)
    lines, cur = [], []

    def walk(el):
        name = _local(el.tag)
        if name == para and cur:
            lines.append("".join(cur))
            cur.clear()
        if name in texts:
            if el.text:
                cur.append(el.text)
            for ch in el:
                cn = _local(ch.tag)
                if cn in tabs:
                    cur.append("\t")
                elif cn in breaks:
                    cur.append("\n")
                elif ch.text:
                    cur.append(ch.text)
                if ch.tail:
                    cur.append(ch.tail)
            return
        if name in tabs:
            cur.append("\t")
        elif name in breaks:
            cur.append("\n")
        for ch in el:
            walk(ch)
        if name == para:
            lines.append("".join(cur))
            cur.clear()

    walk(root)
    if cur:
        lines.append("".join(cur))
    return "\n".join(lines)


def _numbered(names, pattern):
    found = [(int(m.group(1)), n) for n in names for m in [re.search(pattern, n)] if m]
    return [n for _, n in sorted(found)]


def hwpx_text(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    parts = _numbered(z.namelist(), r"^Contents/section(\d+)\.xml$")
    return "\n".join(xml_text(z.read(n)) for n in parts)


def docx_text(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    return xml_text(z.read("word/document.xml"))


def pptx_text(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    out = []
    for i, n in enumerate(_numbered(z.namelist(), r"^ppt/slides/slide(\d+)\.xml$"), 1):
        out.append(f"[슬라이드 {i}]\n" + xml_text(z.read(n)))
    return "\n\n".join(out)


def _cell(v):
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d") if not getattr(v, "hour", 0) else v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).replace("\n", " ")


def _rows_text(title, rows):
    lines, n = [f"[시트: {title}]"], 0
    for row in rows:
        vals = [_cell(v) for v in list(row)[:XLSX_MAX_COLS]]
        while vals and not vals[-1]:
            vals.pop()
        if not vals:
            continue
        n += 1
        if n > XLSX_MAX_ROWS:
            lines.append(f"…({XLSX_MAX_ROWS}행 이후 생략. 원본 파일 참고)")
            break
        lines.append(" | ".join(vals))
    return "\n".join(lines)


def xlsx_text(data):
    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        return "\n\n".join(_rows_text(ws.title, ws.iter_rows(values_only=True)) for ws in wb.worksheets)
    finally:
        wb.close()


def xls_text(data):
    import xlrd  # 선택 설치

    book = xlrd.open_workbook(file_contents=data)
    return "\n\n".join(
        _rows_text(sh.name, (sh.row_values(r) for r in range(sh.nrows))) for sh in book.sheets()
    )


def decode_bytes(data):
    for enc in ("utf-8-sig", "cp949", "utf-16"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", "replace")


def zip_listing(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    names = []
    for info in z.infolist():
        name = info.filename
        if not info.flag_bits & 0x800:  # 윈도우에서 만든 압축파일의 한글 이름
            try:
                name = name.encode("cp437").decode("cp949")
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass
        names.append(f"{name} ({info.file_size:,} bytes)")
    return "압축파일 안의 파일 목록:\n" + "\n".join(names)


def eml_text(data):
    msg = email.message_from_bytes(data, policy=policy.default)
    head = "\n".join(f"{k}: {msg[k]}" for k in ("Date", "From", "To", "Subject") if msg[k])
    return head + "\n\n" + get_body(msg)


def pdf_reader(data):
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(data))
    if reader.is_encrypted:
        reader.decrypt("")
    return reader


def pdf_text(data, max_pages=200):
    reader = pdf_reader(data)
    pages = reader.pages
    texts = []
    for i, page in enumerate(pages):
        if i >= max_pages:
            texts.append(f"…({max_pages}쪽 이후 생략)")
            break
        try:
            texts.append(page.extract_text() or "")
        except Exception:
            texts.append("")
    text = "\n".join(t.strip() for t in texts if t.strip())
    if len(text) < 20 * len(pages):
        note = f"PDF {len(pages)}쪽. 스캔 이미지로 된 PDF라 글자를 뽑지 못한 부분이 있습니다."
    else:
        note = f"PDF {len(pages)}쪽"
    return text, note


EXTRACTORS = {
    "hwp": hwp_text, "hwpx": hwpx_text, "docx": docx_text, "xlsx": xlsx_text, "xls": xls_text,
    "pptx": pptx_text, "text": decode_bytes, "zip": zip_listing, "eml": eml_text,
    "html": lambda d: html_to_text(decode_bytes(d)),
}


def extract(att):
    """첨부 하나에 kind, label, text, note 를 채운다. 실패해도 예외를 내지 않는다."""
    kind = kind_of(att["name"], att["ctype"], att["data"])
    att["kind"], att["label"] = kind, KIND_LABEL[kind]
    att["text"], att["note"] = "", ""
    try:
        if kind == "pdf":
            att["text"], att["note"] = pdf_text(att["data"])
        elif kind in EXTRACTORS:
            att["text"] = EXTRACTORS[kind](att["data"])
            if kind not in ("zip",):
                att["note"] = f"{att['label']}에서 본문 글자만 뽑았습니다. 표·서식은 원본 파일에서 확인하세요."
        elif kind == "image":
            att["note"] = "그림"
        elif kind == "hwp-old":
            att["note"] = "한글 97 이전 형식이라 내용을 읽지 못했습니다. 원본 파일을 여세요."
        else:
            att["note"] = "미리보기를 지원하지 않는 형식입니다. 원본 파일을 여세요."
    except ModuleNotFoundError as e:
        att["note"] = f"{att['label']}을(를) 읽는 프로그램({e.name})이 없어 내용을 뽑지 못했습니다."
    except Exception as e:
        att["note"] = f"내용을 읽지 못했습니다: {e}"
    att["text"] = clean_text(att["text"]).strip()
    return att
