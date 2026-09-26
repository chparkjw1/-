"""이메일(.eml) 폴더를 첨부파일을 뺀 텍스트 파일로 변환한다.

사용법:
    python eml_to_text.py <eml 폴더> [출력 폴더]

출력:
    emails.jsonl  메일 1건당 1줄 (날짜, 보낸사람, 받는사람, 참조, 제목, 첨부파일 이름, 본문)
    emails.txt    같은 내용을 사람이 읽기 좋게 이어 붙인 파일
    본문이 20MB 단위로 나뉘어 emails_001.txt, emails_002.txt ... 로 저장된다.
"""
import email
import html
import json
import re
import sys
from email import policy
from email.utils import parsedate_to_datetime
from pathlib import Path

CHUNK_BYTES = 20 * 1024 * 1024


def html_to_text(s):
    s = re.sub(r"(?is)<(script|style).*?</\1>", "", s)
    s = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", s)
    s = re.sub(r"<[^>]+>", "", s)
    s = html.unescape(s)
    s = re.sub(r"[ \t\xa0]+", " ", s)
    return re.sub(r"\n\s*\n+", "\n\n", s).strip()


def get_body(msg):
    plain, htm = [], []
    for part in msg.walk():
        if part.is_multipart() or part.get_content_disposition() == "attachment":
            continue
        ctype = part.get_content_type()
        if ctype not in ("text/plain", "text/html"):
            continue
        try:
            text = part.get_content()
        except Exception:
            raw = part.get_payload(decode=True) or b""
            text = raw.decode(part.get_content_charset() or "utf-8", "replace")
        (plain if ctype == "text/plain" else htm).append(text)
    if plain and any(p.strip() for p in plain):
        return "\n".join(plain).strip()
    return html_to_text("\n".join(htm))


def attachments(msg):
    names = []
    for part in msg.walk():
        name = part.get_filename()
        if name:
            names.append(name)
    return names


def parse(path):
    with open(path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    try:
        date = parsedate_to_datetime(msg["Date"]).isoformat()
    except Exception:
        date = str(msg["Date"] or "")
    return {
        "file": path.name,
        "date": date,
        "from": str(msg["From"] or ""),
        "to": str(msg["To"] or ""),
        "cc": str(msg["Cc"] or ""),
        "subject": str(msg["Subject"] or ""),
        "attachments": attachments(msg),
        "body": get_body(msg),
    }


def main():
    src = Path(sys.argv[1])
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("eml_text")
    out.mkdir(parents=True, exist_ok=True)
    rows, errors = [], []
    for p in sorted(src.rglob("*.eml")):
        try:
            rows.append(parse(p))
        except Exception as e:
            errors.append(f"{p.name}\t{e}")
    rows.sort(key=lambda r: r["date"])

    with open(out / "emails.jsonl", "w", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

    chunk, size, n = [], 0, 1
    def flush():
        nonlocal chunk, size, n
        if chunk:
            (out / f"emails_{n:03d}.txt").write_text("".join(chunk), encoding="utf-8")
            chunk, size, n = [], 0, n + 1
    for r in rows:
        block = (
            f"===== {r['date']} | {r['subject']}\n"
            f"보낸사람: {r['from']}\n받는사람: {r['to']}\n"
            + (f"참조: {r['cc']}\n" if r["cc"] else "")
            + (f"첨부: {', '.join(r['attachments'])}\n" if r["attachments"] else "")
            + f"파일: {r['file']}\n\n{r['body']}\n\n"
        )
        b = len(block.encode("utf-8"))
        if size + b > CHUNK_BYTES:
            flush()
        chunk.append(block)
        size += b
    flush()

    if errors:
        (out / "errors.txt").write_text("\n".join(errors), encoding="utf-8")
    print(f"변환 {len(rows)}건, 실패 {len(errors)}건 -> {out.resolve()}")


if __name__ == "__main__":
    main()
