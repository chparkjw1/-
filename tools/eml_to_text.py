"""이메일(.eml) 폴더를 본문과 첨부 글자를 모은 텍스트 파일로 변환한다. (분석용)

사용법:
    pip install pypdf olefile openpyxl
    python eml_to_text.py <eml 폴더> [출력 폴더] [--rules 규칙.json] [--list 목록]

출력 (출력 폴더 기본값: 분석용_텍스트):
    emails.jsonl         메일 1건당 1줄 (날짜, 보낸사람, 받는사람, 참조, 제목, 본문, 첨부 이름·글자)
    emails_001.txt ...   같은 내용을 사람이 읽기 좋게 이어 붙인 파일 (약 400KB씩 나눔)
    오류.txt             읽지 못한 메일 (있을 때만)
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from attachments import extract  # noqa: E402
from mailcommon import parse, parse_args, select_files  # noqa: E402

CHUNK_BYTES = 400 * 1024       # 드라이브에서 한 번에 읽기 좋은 크기
MAX_ATT_TEXT = 30000           # 첨부 하나에서 텍스트 파일에 넣는 최대 글자 수


def cut(text, limit, what):
    if len(text) <= limit:
        return text
    return text[:limit] + f"\n…({what}이 길어 이하 생략, 전체 {len(text):,}자)"


def email_block(r):
    lines = [
        f"===== {r['date']} | {r['subject']}",
        f"보낸사람: {r['from']}",
        f"받는사람: {r['to']}",
    ]
    if r["cc"]:
        lines.append(f"참조: {r['cc']}")
    if r["attachments"]:
        lines.append("첨부: " + ", ".join(a["name"] for a in r["attachments"]))
    lines.append(f"파일: {r['file']}")
    lines += ["", r["body"] or "(본문 없음)", ""]
    for j, a in enumerate(r["attachments"], 1):
        lines.append(f"--- 첨부 {j}: {a['name']} ({a.get('note') or a.get('label', '')}) ---")
        if a.get("text"):
            lines.append(cut(a["text"], MAX_ATT_TEXT, "첨부 내용"))
        lines.append("")
    return "\n".join(lines) + "\n"


def write_text_outputs(rows, out):
    """이미 읽은 메일(첨부 글자 포함)을 jsonl과 나눈 txt로 저장한다."""
    out.mkdir(parents=True, exist_ok=True)
    rows = sorted(rows, key=lambda r: r["date"])
    with open(out / "emails.jsonl", "w", encoding="utf-8") as f:
        for r in rows:
            rec = {k: r[k] for k in ("file", "date", "from", "to", "cc", "subject", "body")}
            rec["attachments"] = [
                {"name": a["name"], "kind": a.get("kind"), "note": a.get("note"), "text": a.get("text", "")}
                for a in r["attachments"]
            ]
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    chunk, size, n = [], 0, 1
    for r in rows:
        block = email_block(r)
        b = len(block.encode("utf-8"))
        if chunk and size + b > CHUNK_BYTES:
            (out / f"emails_{n:03d}.txt").write_text("".join(chunk), encoding="utf-8")
            chunk, size, n = [], 0, n + 1
        chunk.append(block)
        size += b
    if chunk:
        (out / f"emails_{n:03d}.txt").write_text("".join(chunk), encoding="utf-8")
    return n if chunk else n - 1


def main():
    src, out, opts = parse_args(sys.argv[1:], "분석용_텍스트")
    files, missing = select_files(src, opts)
    rows, errors = [], []
    for p in files:
        try:
            r = parse(p)
            for a in r["attachments"]:
                extract(a)
            rows.append(r)
        except Exception as e:
            errors.append(f"{p.name}\t{e}")
    parts = write_text_outputs(rows, out)
    if missing:
        errors += [f"{m}\t목록에 있으나 폴더에서 찾지 못함" for m in missing]
    if errors:
        (out / "오류.txt").write_text("\n".join(errors), encoding="utf-8")
    print(f"변환 {len(rows)}건, 텍스트 파일 {parts}개, 문제 {len(errors)}건 -> {out.resolve()}")


if __name__ == "__main__":
    main()
