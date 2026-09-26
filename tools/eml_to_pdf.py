"""이메일(.eml) 폴더를 읽기 쉬운 PDF로 변환한다.

사용법:
    pip install reportlab
    python eml_to_pdf.py <eml 폴더> [출력 폴더] [--by month|year] [--list 목록파일]

    --list: 변환할 메일 파일명 목록(첫 열이 .eml 파일명인 TSV/TXT). 지정하면 목록에 있는 메일만 변환한다.

출력 (출력 폴더 기본값: 메일_PDF):
    PDF/2023-01.pdf ...   월(또는 연)별 PDF. 첫 장은 목차, 이후 메일 1건씩 새 페이지에서 시작.
                          PDF 책갈피(북마크)로 메일 제목을 바로 찾아갈 수 있다.
    첨부파일/<날짜_제목>/  메일마다 첨부파일 원본을 저장한 폴더
    목록.csv              전체 메일 목록 (엑셀에서 바로 열림)
"""
import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import (
    Flowable, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

sys.path.insert(0, str(Path(__file__).parent))
from eml_to_text import parse  # noqa: E402

import email
from email import policy

FONT = "HYGothic-Medium"
pdfmetrics.registerFont(UnicodeCIDFont(FONT))

BASE = dict(fontName=FONT, wordWrap="CJK", alignment=TA_LEFT)
S_TITLE = ParagraphStyle("title", fontSize=15, leading=21, spaceAfter=6, **BASE)
S_SUBJ = ParagraphStyle("subj", fontSize=12.5, leading=18, spaceAfter=4, **BASE)
S_LABEL = ParagraphStyle("label", fontSize=8.5, leading=12, textColor=colors.HexColor("#555555"), **BASE)
S_CELL = ParagraphStyle("cell", fontSize=8.5, leading=12, **BASE)
S_BODY = ParagraphStyle("body", fontSize=9.5, leading=14.5, **BASE)
S_TOC = ParagraphStyle("toc", fontSize=8, leading=11, **BASE)

MAX_BODY_CHARS = 60000  # 한 메일 본문이 이보다 길면 잘라서 표시 (원본은 eml에 그대로 있음)


def esc(s):
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def safe_name(s, limit=60):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", s or "").strip(" ._")
    return (s or "제목없음")[:limit]


class Bookmark(Flowable):
    """PDF 책갈피를 남기는 보이지 않는 요소."""

    def __init__(self, key, title):
        super().__init__()
        self.key, self.title = key, title
        self.width = self.height = 0

    def draw(self):
        c = self.canv
        c.bookmarkPage(self.key)
        c.addOutlineEntry(self.title, self.key, level=0, closed=True)


def save_attachments(path, folder):
    with open(path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    saved = []
    for part in msg.walk():
        name = part.get_filename()
        if not name:
            continue
        data = part.get_payload(decode=True)
        if not data:
            continue
        folder.mkdir(parents=True, exist_ok=True)
        target = folder / safe_name(name, 120)
        i = 1
        while target.exists():
            target = folder / f"{target.stem}_{i}{target.suffix}"
            i += 1
        target.write_bytes(data)
        saved.append(target.name)
    return saved


def header_table(r, att_dir):
    rows = [
        ("날짜", r["date"].replace("T", " ")[:19]),
        ("보낸사람", r["from"]),
        ("받는사람", r["to"]),
    ]
    if r["cc"]:
        rows.append(("참조", r["cc"]))
    if r["attachments"]:
        rows.append(("첨부", "<br/>".join(esc(a) for a in r["attachments"])
                     + f"<br/><font color='#777777'>저장 위치: {esc(att_dir)}</font>"))
    rows.append(("원본 파일", r["file"]))
    data = [[Paragraph(k, S_LABEL), Paragraph(v if k == "첨부" else esc(v), S_CELL)] for k, v in rows]
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


def body_flowables(text):
    text = (text or "").strip() or "(본문 없음)"
    if len(text) > MAX_BODY_CHARS:
        text = text[:MAX_BODY_CHARS] + "\n\n…(본문이 길어 이하 생략. 원본 eml 참고)"
    out = []
    for para in re.split(r"\n\s*\n", text):
        lines = [esc(l) for l in para.splitlines()]
        out.append(Paragraph("<br/>".join(lines), S_BODY))
        out.append(Spacer(1, 4))
    return out


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont(FONT, 7.5)
    canvas.setFillColor(colors.HexColor("#888888"))
    canvas.drawString(15 * mm, 9 * mm, doc.title)
    canvas.drawRightString(A4[0] - 15 * mm, 9 * mm, f"{doc.page}쪽")
    canvas.restoreState()


def build_pdf(group, rows, pdf_path):
    doc = SimpleDocTemplate(
        str(pdf_path), pagesize=A4, title=f"이메일 {group}",
        leftMargin=15 * mm, rightMargin=15 * mm, topMargin=15 * mm, bottomMargin=15 * mm,
    )
    story = [Paragraph(f"이메일 {esc(group)} — {len(rows)}건", S_TITLE)]
    toc = [[Paragraph("번호", S_LABEL), Paragraph("날짜", S_LABEL),
            Paragraph("보낸사람", S_LABEL), Paragraph("제목", S_LABEL)]]
    for i, r in enumerate(rows, 1):
        sender = re.sub(r"\s*<.*?>", "", r["from"]) or r["from"]
        toc.append([Paragraph(str(i), S_TOC), Paragraph(r["date"][:10], S_TOC),
                    Paragraph(esc(sender[:30]), S_TOC),
                    Paragraph(f"<a href='#m{i}' color='#1a4fa0'>{esc(r['subject'] or '(제목 없음)')}</a>", S_TOC)])
    t = Table(toc, colWidths=[12 * mm, 20 * mm, 38 * mm, None], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F4F7")),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(t)

    for i, r in enumerate(rows, 1):
        story.append(PageBreak())
        subj = r["subject"] or "(제목 없음)"
        story.append(Bookmark(f"m{i}", f"{r['date'][:10]} {subj}"[:120]))
        story.append(Paragraph(f"<a name='m{i}'/>{i}. {esc(subj)}", S_SUBJ))
        story.append(header_table(r, r["_att_dir"]))
        story.append(Spacer(1, 8))
        story.extend(body_flowables(r["body"]))

    doc.build(story, onFirstPage=footer, onLaterPages=footer)


def main():
    argv = sys.argv[1:]
    opts = {}
    for flag in ("--by", "--list"):
        if flag in argv:
            i = argv.index(flag)
            opts[flag] = argv[i + 1]
            del argv[i:i + 2]
    args = argv
    by_year = opts.get("--by") == "year"
    only = None
    if "--list" in opts:
        with open(opts["--list"], encoding="utf-8-sig") as f:
            only = {line.split("\t")[0].strip() for line in f if line.strip()}
    src = Path(args[0])
    out = Path(args[1]) if len(args) > 1 else Path("메일_PDF")
    (out / "PDF").mkdir(parents=True, exist_ok=True)

    rows, errors = [], []
    for p in sorted(src.rglob("*.eml")):
        if only is not None and p.name not in only:
            continue
        try:
            r = parse(p)
            r["_path"] = p
            rows.append(r)
        except Exception as e:
            errors.append(f"{p.name}\t{e}")
    rows.sort(key=lambda r: r["date"])

    groups = defaultdict(list)
    for r in rows:
        key = r["date"][:4] if by_year else r["date"][:7]
        if not re.match(r"\d{4}", key):
            key = "날짜미상"
        groups[key].append(r)

    with open(out / "목록.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["PDF", "번호", "날짜", "보낸사람", "받는사람", "제목", "첨부파일", "원본파일"])
        for g, items in sorted(groups.items()):
            for i, r in enumerate(items, 1):
                att_dir = out / "첨부파일" / safe_name(f"{r['date'][:10]}_{r['subject']}")
                r["_att_dir"] = str(att_dir.relative_to(out)) if r["attachments"] else ""
                if r["attachments"]:
                    try:
                        save_attachments(r["_path"], att_dir)
                    except Exception as e:
                        errors.append(f"{r['file']}\t첨부 저장 실패: {e}")
                w.writerow([f"{g}.pdf", i, r["date"][:19].replace("T", " "), r["from"], r["to"],
                            r["subject"], " | ".join(r["attachments"]), r["file"]])

    for g, items in sorted(groups.items()):
        try:
            build_pdf(g, items, out / "PDF" / f"{g}.pdf")
            print(f"{g}.pdf  {len(items)}건")
        except Exception as e:
            errors.append(f"{g}.pdf 생성 실패: {e}")

    if errors:
        (out / "오류.txt").write_text("\n".join(errors), encoding="utf-8")
    if only is not None:
        missing = only - {r["file"] for r in rows} - {"파일명"}
        if missing:
            (out / "목록에있으나_없는파일.txt").write_text("\n".join(sorted(missing)), encoding="utf-8")
            print(f"목록에 있으나 폴더에서 찾지 못한 파일 {len(missing)}건 -> 목록에있으나_없는파일.txt")
    print(f"완료: 메일 {len(rows)}건, PDF {len(groups)}개, 오류 {len(errors)}건 -> {out.resolve()}")


if __name__ == "__main__":
    main()
