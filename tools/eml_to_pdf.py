"""이메일(.eml) 폴더를 읽기 쉬운 PDF로 변환한다. 첨부파일도 함께 꺼내고 PDF 안에 펼쳐 보여 준다.

사용법:
    pip install reportlab pypdf olefile openpyxl pillow
    python eml_to_pdf.py <eml 폴더> [출력 폴더] [--rules 규칙.json] [--list 목록] [--by month|year] [--hwp-pdf]

    --rules    파일명 규칙(JSON)에 맞는 메일만 변환 (mailcommon.load_rules 참고)
    --list     목록 파일(첫 열이 .eml 파일명)에 있는 메일만 변환
    --by year  월별 대신 연도별 PDF로 묶기
    --hwp-pdf  (윈도우 + 한글 설치 PC) 한글 첨부를 한글 프로그램으로 PDF로 바꿔 원래 모양대로 붙인다.
               실패하면 자동으로 본문 글자만 뽑아 넣는다.

출력 (출력 폴더 기본값: 메일_PDF):
    PDF/2023-01.pdf     월별 PDF. 첫 장은 목차(제목을 누르면 해당 메일로 이동), 이후 메일 1건씩.
                        메일마다 본문 뒤에 첨부 내용이 이어진다.
                          - PDF 첨부: 원본 페이지를 그대로 붙임
                          - 그림 첨부: 그림을 그대로 넣음
                          - 한글·워드·엑셀·파워포인트·텍스트 첨부: 본문 글자를 뽑아 넣음
                        왼쪽 책갈피에서 메일·첨부 제목으로 바로 이동할 수 있다.
    첨부파일/<날짜_제목>/  메일마다 첨부 원본 파일
    분석용_텍스트/      본문과 첨부 글자를 모은 텍스트 (emails_001.txt ...)
    목록.csv            전체 메일 목록 (엑셀에서 바로 열림)
    오류.txt            변환 중 문제가 생긴 메일 (있을 때만)
"""
import csv
import io
import os
import re
import sys
import tempfile
from collections import defaultdict
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Flowable, Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

sys.path.insert(0, str(Path(__file__).parent))
from attachments import extract, pdf_reader  # noqa: E402
from eml_to_text import write_text_outputs  # noqa: E402
from mailcommon import parse, parse_args, safe_name, select_files  # noqa: E402

MAX_BODY_CHARS = 60000   # 본문이 이보다 길면 PDF에는 잘라서 표시 (원본 eml·분석용 텍스트에는 전부 있음)
MAX_ATT_CHARS = 30000    # 첨부에서 뽑은 글자를 PDF에 넣는 최대 길이
MAX_APPEND_PAGES = 300   # 이보다 긴 PDF 첨부는 붙이지 않고 원본 파일로 안내
IMG_MAX_PX = 1600        # 사진은 이 크기로 줄여 넣는다 (원본은 첨부파일 폴더에 그대로)
FRAME_W = A4[0] - 30 * mm
FRAME_H = A4[1] - 40 * mm


# ---------------------------------------------------------------- 글꼴
def register_font():
    """PC에 있는 한글 글꼴을 PDF에 넣는다. 없으면 PDF 표준 한글 글꼴을 쓴다."""
    home = Path.home()
    candidates = [
        "C:/Windows/Fonts/malgun.ttf",
        "C:/Windows/Fonts/NanumGothic.ttf",
        str(home / "AppData/Local/Microsoft/Windows/Fonts/NanumGothic.ttf"),
        "/Library/Fonts/NanumGothic.ttf",
        str(home / "Library/Fonts/NanumGothic.ttf"),
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
        "C:/Windows/Fonts/gulim.ttc",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("Korean", path, subfontIndex=0))
                return "Korean"
            except Exception:
                continue
    pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
    return "HYGothic-Medium"


FONT = register_font()
BASE = dict(fontName=FONT, wordWrap="CJK", alignment=TA_LEFT)
S_TITLE = ParagraphStyle("title", fontSize=15, leading=21, spaceAfter=6, **BASE)
S_SUBJ = ParagraphStyle("subj", fontSize=12.5, leading=18, spaceAfter=4, **BASE)
S_ATT = ParagraphStyle("att", fontSize=10.5, leading=15, spaceBefore=10, spaceAfter=4,
                       textColor=colors.HexColor("#1a3f7a"), **BASE)
S_NOTE = ParagraphStyle("note", fontSize=8.5, leading=12, spaceAfter=4,
                        textColor=colors.HexColor("#666666"), **BASE)
S_LABEL = ParagraphStyle("label", fontSize=8.5, leading=12, textColor=colors.HexColor("#555555"), **BASE)
S_CELL = ParagraphStyle("cell", fontSize=8.5, leading=12, **BASE)
S_BODY = ParagraphStyle("body", fontSize=9.5, leading=14.5, **BASE)
S_ATTBODY = ParagraphStyle("attbody", fontSize=9, leading=13.5, **BASE)
S_TOC = ParagraphStyle("toc", fontSize=8, leading=11, **BASE)


def esc(s):
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def text_flowables(text, style, limit, cut_note):
    text = (text or "").strip()
    if len(text) > limit:
        text = text[:limit] + f"\n\n…({cut_note})"
    out = []
    for para in re.split(r"\n\s*\n", text):
        if para.strip():
            out.append(Paragraph("<br/>".join(esc(l) for l in para.splitlines()), style))
            out.append(Spacer(1, 3))
    return out


class Marker(Flowable):
    """보이지 않는 표시. 그려지는 쪽 번호를 기록해 책갈피에 쓴다."""

    def __init__(self, store, key):
        super().__init__()
        self.store, self.key = store, key
        self.width = self.height = 0

    def draw(self):
        self.store[self.key] = self.canv.getPageNumber()


# ---------------------------------------------------------------- 첨부 → 그림/PDF
def image_flowable(data):
    from PIL import Image as PILImage, ImageOps

    im = PILImage.open(io.BytesIO(data))
    im.seek(0)
    im = ImageOps.exif_transpose(im)
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    im.thumbnail((IMG_MAX_PX, IMG_MAX_PX))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=85)
    buf.seek(0)
    w, h = im.size
    scale = min(FRAME_W / w, (FRAME_H - 45 * mm) / h, 1.5)
    return Image(buf, width=w * scale, height=h * scale)


_HWP = None


def hwp_to_pdf(data, suffix):
    """윈도우의 한글 프로그램으로 한글 문서를 PDF로 바꾼다. (--hwp-pdf 사용 시)"""
    global _HWP
    import win32com.client as win32

    if _HWP is None:
        _HWP = win32.Dispatch("HWPFrame.HwpObject")
        try:
            _HWP.XHwpWindows.Item(0).Visible = False
        except Exception:
            pass
        try:
            _HWP.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / f"att{suffix}"
        dst = Path(tmp) / "att.pdf"
        src.write_bytes(data)
        if not _HWP.Open(str(src), "", "forceopen:true"):
            raise RuntimeError("한글에서 파일을 열지 못했습니다")
        _HWP.SaveAs(str(dst), "PDF", "")
        _HWP.Clear(1)
        return dst.read_bytes()


def close_hwp():
    if _HWP is not None:
        try:
            _HWP.Quit()
        except Exception:
            pass


# ---------------------------------------------------------------- 메일 1건 → PDF 조각
def header_table(r):
    rows = [("날짜", esc(r["date"].replace("T", " ")[:19])), ("보낸사람", esc(r["from"])), ("받는사람", esc(r["to"]))]
    if r["cc"]:
        rows.append(("참조", esc(r["cc"])))
    if r["attachments"]:
        names = "<br/>".join(f"{i}. {esc(a['name'])}" for i, a in enumerate(r["attachments"], 1))
        rows.append(("첨부", names + f"<br/><font color='#777777'>원본 저장 위치: {esc(r['att_dir'])}</font>"))
    rows.append(("원본 파일", esc(r["file"])))
    data = [[Paragraph(k, S_LABEL), Paragraph(v, S_CELL)] for k, v in rows]
    t = Table(data, colWidths=[22 * mm, None])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F2F4F7")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#BBBBBB")),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def email_piece(r, idx, total, group, marks):
    """메일 1건(본문 + 그림·글자 첨부)을 PDF로 만든다. PDF 첨부는 따로 뒤에 붙인다."""
    subj = r["subject"] or "(제목 없음)"
    story = [Marker(marks, "start"), Paragraph(f"{idx}. {esc(subj)}", S_SUBJ), header_table(r), Spacer(1, 8)]
    story += text_flowables(r["body"] or "(본문 없음)", S_BODY, MAX_BODY_CHARS,
                            "본문이 길어 이하 생략. 분석용 텍스트·원본 eml 참고")
    for j, a in enumerate(r["attachments"], 1):
        head = [Marker(marks, f"att{j}"), Paragraph(f"첨부 {j}. {esc(a['name'])}", S_ATT)]
        if a.get("pdf_bytes"):
            head.append(Paragraph(f"{esc(a['note'])} — 원본 페이지를 이 메일 뒤에 붙였습니다.", S_NOTE))
            story.append(KeepTogether(head))
        elif a["kind"] == "image":
            try:
                head.append(image_flowable(a["data"]))
            except Exception as e:
                head.append(Paragraph(f"그림을 넣지 못했습니다({esc(str(e))}). 원본 파일을 여세요.", S_NOTE))
            story.append(KeepTogether(head))
        else:
            if a["note"]:
                head.append(Paragraph(esc(a["note"]), S_NOTE))
            body = text_flowables(a["text"], S_ATTBODY, MAX_ATT_CHARS,
                                  "첨부 내용이 길어 이하 생략. 원본 파일 참고") if a["text"] else []
            story.append(KeepTogether(head + body[:1]))
            story += body[1:]

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(FONT, 7.5)
        canvas.setFillColor(colors.HexColor("#888888"))
        canvas.drawString(15 * mm, 9 * mm, f"이메일 {group} · {idx}/{total} · {r['date'][:10]}")
        canvas.drawRightString(A4[0] - 15 * mm, 9 * mm, subj[:40])
        canvas.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=15 * mm, bottomMargin=15 * mm, title=subj)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


def fallback_piece(r, idx, total, group, err):
    """모양을 갖춘 조각을 만들다 실패하면, 최소한 본문만이라도 담는다."""
    story = [Paragraph(f"{idx}. {esc(r['subject'] or '(제목 없음)')}", S_SUBJ),
             Paragraph(esc(f"※ 이 메일은 서식을 갖춰 만들지 못해 간단히 표시합니다 ({err})"), S_NOTE),
             Paragraph(esc(f"날짜: {r['date']}  보낸사람: {r['from']}"), S_CELL), Spacer(1, 6)]
    story += text_flowables(r["body"], S_BODY, MAX_BODY_CHARS, "이하 생략")
    buf = io.BytesIO()
    SimpleDocTemplate(buf, pagesize=A4).build(story)
    return buf.getvalue()


# ---------------------------------------------------------------- 월별 PDF 합치기
def toc_pdf(group, rows, start_pages):
    head = [Paragraph(x, S_LABEL) for x in ("번호", "날짜", "보낸사람", "제목", "첨부", "쪽")]
    data = [head]
    for i, r in enumerate(rows, 1):
        sender = re.sub(r"\s*<.*?>", "", r["from"]).strip().strip('"') or r["from"]
        title = esc(r["subject"] or "(제목 없음)")
        data.append([
            Paragraph(str(i), S_TOC), Paragraph(r["date"][:10], S_TOC), Paragraph(esc(sender[:30]), S_TOC),
            Paragraph(f"<a href='goto:{i}' color='#1a4fa0'>{title}</a>", S_TOC),
            Paragraph(str(len(r["attachments"])) if r["attachments"] else "", S_TOC),
            Paragraph(str(start_pages[i - 1]) if start_pages else "", S_TOC),
        ])
    t = Table(data, colWidths=[12 * mm, 22 * mm, 32 * mm, None, 12 * mm, 12 * mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F4F7")),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story = [Paragraph(f"이메일 {esc(group)} — {len(rows)}건", S_TITLE),
             Paragraph("제목을 누르면 해당 메일로 이동합니다. 왼쪽 책갈피에서도 찾을 수 있습니다.", S_NOTE), t]
    buf = io.BytesIO()
    SimpleDocTemplate(buf, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
                      topMargin=15 * mm, bottomMargin=15 * mm, title=f"이메일 {group}").build(story)
    return buf.getvalue()


def build_group_pdf(group, rows, pdf_path, errors):
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import ArrayObject, NameObject

    pieces = []  # (메일 조각 PDF, 첨부 PDF 목록, 쪽 표시)
    for i, r in enumerate(rows, 1):
        marks = {}
        try:
            piece = email_piece(r, i, len(rows), group, marks)
        except Exception as e:
            errors.append(f"{r['file']}\tPDF 서식 만들기 실패, 간단히 표시함: {e}")
            piece, marks = fallback_piece(r, i, len(rows), group, e), {"start": 1}
        appended = [(j, a) for j, a in enumerate(r["attachments"], 1) if a.get("pdf_bytes")]
        pieces.append((PdfReader(io.BytesIO(piece)), appended, marks))

    def pages_of(item):
        reader, appended, _ = item
        return len(reader.pages) + sum(len(a["pdf_reader"].pages) for _, a in appended)

    toc_pages = len(PdfReader(io.BytesIO(toc_pdf(group, rows, None))).pages)
    starts, cur = [], toc_pages + 1
    for item in pieces:
        starts.append(cur)
        cur += pages_of(item)
    toc = PdfReader(io.BytesIO(toc_pdf(group, rows, starts)))

    writer = PdfWriter()
    writer.append(toc, import_outline=False)
    for i, (r, (reader, appended, marks)) in enumerate(zip(rows, pieces), 1):
        base = len(writer.pages)
        writer.append(reader, import_outline=False)
        att_pages = {j: base + marks[f"att{j}"] - 1 for j in range(1, len(r["attachments"]) + 1)
                     if f"att{j}" in marks}
        for j, a in appended:
            at = len(writer.pages)
            try:
                writer.append(a["pdf_reader"], import_outline=False)
                att_pages[j] = at
            except Exception as e:
                errors.append(f"{r['file']}\t첨부 PDF 붙이기 실패({a['name']}): {e}")
        parent = writer.add_outline_item(f"{i}. {r['date'][:10]} {r['subject'] or '(제목 없음)'}"[:120], base)
        for j, a in enumerate(r["attachments"], 1):
            if j in att_pages:
                writer.add_outline_item(f"첨부 {j}. {a['name']}"[:120], att_pages[j], parent=parent)

    # 목차의 제목 링크(goto:번호)를 해당 메일 첫 쪽으로 연결한다.
    for p in range(toc_pages):
        for annot in writer.pages[p].get("/Annots", []) or []:
            a = annot.get_object()
            action = a.get("/A")
            uri = str(action.get_object().get("/URI", "")) if action else ""
            if uri.startswith("goto:"):
                target = starts[int(uri[5:]) - 1] - 1
                if target < len(writer.pages):
                    a[NameObject("/Dest")] = ArrayObject([writer.pages[target].indirect_reference, NameObject("/Fit")])
                    del a["/A"]
    writer.page_mode = "/UseOutlines"
    with open(pdf_path, "wb") as f:
        writer.write(f)


# ---------------------------------------------------------------- 실행
def prepare_attachments(r, out, used_dirs, use_hwp_pdf, errors):
    """첨부를 원본 파일로 저장하고, PDF에 넣을 형태(글자·그림·PDF 페이지)를 준비한다."""
    r["att_dir"] = ""
    if not r["attachments"]:
        return
    base = safe_name(f"{r['date'][:10]}_{r['subject']}", 50)
    folder, n = base, 2
    while folder in used_dirs:
        folder, n = f"{base}_{n}", n + 1
    used_dirs.add(folder)
    target_dir = out / "첨부파일" / folder
    target_dir.mkdir(parents=True, exist_ok=True)
    r["att_dir"] = f"첨부파일/{folder}"
    for a in r["attachments"]:
        name = safe_name(a["name"], 80)
        dest, k = target_dir / name, 1
        while dest.exists():
            stem, suf = os.path.splitext(name)
            dest, k = target_dir / f"{stem}_{k}{suf}", k + 1
        try:
            dest.write_bytes(a["data"])
        except Exception as e:
            errors.append(f"{r['file']}\t첨부 저장 실패({a['name']}): {e}")
        extract(a)
        pdf_bytes = None
        if a["kind"] == "pdf":
            pdf_bytes = a["data"]
        elif use_hwp_pdf and a["kind"] in ("hwp", "hwpx"):
            try:
                pdf_bytes = hwp_to_pdf(a["data"], "." + a["kind"])
                a["note"] = "한글 문서를 PDF로 바꿔 붙였습니다"
            except Exception as e:
                errors.append(f"{r['file']}\t한글→PDF 변환 실패, 글자만 넣음({a['name']}): {e}")
        if pdf_bytes:
            try:
                reader = pdf_reader(pdf_bytes)
                pages = len(reader.pages)
                if pages <= MAX_APPEND_PAGES:
                    a["pdf_bytes"], a["pdf_reader"] = pdf_bytes, reader
                    if a["kind"] == "pdf" and not a["note"]:
                        a["note"] = f"PDF {pages}쪽"
                else:
                    a["note"] = f"PDF {pages}쪽이라 붙이지 않았습니다. 원본 파일을 여세요."
            except Exception as e:
                a["note"] = f"PDF를 열지 못했습니다({e}). 원본 파일을 여세요."


def main():
    src, out, opts = parse_args(sys.argv[1:], "메일_PDF")
    by_year = opts.get("--by") == "year"
    (out / "PDF").mkdir(parents=True, exist_ok=True)

    files, missing = select_files(src, opts)
    print(f"변환할 메일 {len(files)}건")
    rows, errors, used_dirs = [], [], set()
    for n, p in enumerate(files, 1):
        try:
            r = parse(p)
        except Exception as e:
            errors.append(f"{p.name}\t메일 읽기 실패: {e}")
            continue
        prepare_attachments(r, out, used_dirs, opts.get("--hwp-pdf"), errors)
        rows.append(r)
        if n % 50 == 0:
            print(f"  메일 읽는 중 {n}/{len(files)}")
    close_hwp()
    rows.sort(key=lambda r: r["date"])

    groups = defaultdict(list)
    for r in rows:
        key = r["date"][:4] if by_year else r["date"][:7]
        groups[key if re.match(r"\d{4}", key) else "날짜미상"].append(r)

    for g, items in sorted(groups.items()):
        try:
            build_group_pdf(g, items, out / "PDF" / f"{g}.pdf", errors)
            print(f"{g}.pdf  {len(items)}건")
        except Exception as e:
            errors.append(f"{g}.pdf\tPDF 만들기 실패: {e}")

    with open(out / "목록.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["PDF", "번호", "날짜", "보낸사람", "받는사람", "제목", "첨부 수", "첨부파일", "첨부 저장 위치", "원본파일"])
        for g, items in sorted(groups.items()):
            for i, r in enumerate(items, 1):
                w.writerow([f"{g}.pdf", i, r["date"][:19].replace("T", " "), r["from"], r["to"], r["subject"],
                            len(r["attachments"]), " | ".join(a["name"] for a in r["attachments"]),
                            r["att_dir"], r["file"]])

    write_text_outputs(rows, out / "분석용_텍스트")

    if missing:
        (out / "목록에있으나_없는파일.txt").write_text("\n".join(missing), encoding="utf-8")
        print(f"목록에 있으나 폴더에서 찾지 못한 파일 {len(missing)}건 -> 목록에있으나_없는파일.txt")
    if errors:
        (out / "오류.txt").write_text("\n".join(errors), encoding="utf-8")
    n_att = sum(len(r["attachments"]) for r in rows)
    print(f"완료: 메일 {len(rows)}건, 첨부 {n_att}개, PDF {len(groups)}개, 문제 {len(errors)}건 -> {out.resolve()}")


if __name__ == "__main__":
    main()
