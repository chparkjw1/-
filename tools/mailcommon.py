"""eml_to_pdf.py / eml_to_text.py 가 함께 쓰는 공통 기능: 메일 읽기, 메일 고르기, 파일명 정리."""
import email
import html
import json
import re
import unicodedata
from email import policy
from email.utils import parsedate_to_datetime
from pathlib import Path

MAX_HEADER_CHARS = 1000  # 받는사람·참조가 수백 명인 전체메일은 앞부분만 표시


def nfc(s):
    return unicodedata.normalize("NFC", s or "")


def clean_text(s):
    """PDF·텍스트에 넣을 수 없는 제어문자와 깨진 문자를 정리한다."""
    s = (s or "").encode("utf-8", "replace").decode("utf-8")
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", s)


def safe_name(s, limit=60):
    """윈도우에서도 쓸 수 있는 파일·폴더 이름으로 바꾼다."""
    s = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "_", nfc(s)).strip(" ._")
    return (s or "제목없음")[:limit].strip(" ._") or "제목없음"


def fix_filename(name):
    """메일 프로그램이 한글 파일명을 인코딩 없이 보낸 경우 깨진 글자를 되살린다."""
    if not name:
        return name
    if re.search(r"[\udc80-\udcff]", name):
        raw = name.encode("utf-8", "surrogateescape")
        for enc in ("utf-8", "cp949"):
            try:
                return raw.decode(enc)
            except UnicodeDecodeError:
                pass
        return raw.decode("cp949", "replace")
    return name


def raw_filename(part):
    """헤더를 가공하기 전 원문에서 파일명을 꺼내 cp949/utf-8로 풀어 본다."""
    for key, value in part.raw_items():
        if key.lower() not in ("content-disposition", "content-type"):
            continue
        m = re.search(r'(?i)\b(?:file)?name\s*=\s*"?([^";\r\n]+)', value)
        if m:
            raw = m.group(1).strip()
            if re.search(r"[\udc80-\udcff]", raw):
                return fix_filename(raw)
    return None


def html_to_text(s):
    s = re.sub(r"(?is)<(script|style|head).*?</\1>", "", s)
    s = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>|</li>|</h\d>", "\n", s)
    s = re.sub(r"(?i)</td>|</th>", "\t", s)
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
        if ctype not in ("text/plain", "text/html") or part.get_filename():
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


def get_attachments(msg):
    """첨부파일과 본문에 끼워 넣은 이미지를 (이름, 형식, 내용) 목록으로 돌려준다."""
    out, k = [], 0
    for part in msg.walk():
        if part.is_multipart():
            continue
        name = fix_filename(part.get_filename())
        if not name or "\ufffd" in name:
            name = raw_filename(part) or name
        ctype = part.get_content_type()
        disp = part.get_content_disposition()
        if not name:
            if part.get_content_maintype() == "image":
                k += 1
                name = f"본문이미지_{k}.{part.get_content_subtype()}"
            elif disp == "attachment":
                k += 1
                name = f"첨부_{k}"
            else:
                continue
        data = part.get_payload(decode=True)
        if not data:
            continue
        if part.get_content_maintype() == "image" and len(data) < 1024:
            continue  # 서명 로고·빈 이미지 같은 아주 작은 그림은 뺀다
        out.append({"name": nfc(name), "ctype": ctype, "data": data})
    return out


def header(msg, key):
    try:
        v = str(msg[key] or "")
    except Exception:
        v = ""
    v = clean_text(v)
    return v if len(v) <= MAX_HEADER_CHARS else v[:MAX_HEADER_CHARS] + " …(이하 생략)"


def parse(path):
    """eml 파일 하나를 읽어 날짜·보낸사람·받는사람·제목·본문·첨부를 돌려준다."""
    with open(path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    try:
        date = parsedate_to_datetime(msg["Date"]).isoformat()
    except Exception:
        date = str(msg["Date"] or "")
    return {
        "file": nfc(Path(path).name),
        "path": Path(path),
        "date": date,
        "from": header(msg, "From"),
        "to": header(msg, "To"),
        "cc": header(msg, "Cc"),
        "subject": header(msg, "Subject"),
        "body": clean_text(get_body(msg)),
        "attachments": get_attachments(msg),
    }


def load_rules(path):
    """파일명으로 메일을 고르는 규칙(JSON)을 읽는다.

    {"include": "정규식", "senders": ["보낸사람", ...], "exclude": "정규식"}
    파일명이 include에 맞거나, 파일명 끝의 _'보낸사람'이 senders에 있으면 고른다.
    단, exclude에 맞으면 뺀다.
    """
    with open(path, encoding="utf-8-sig") as f:
        cfg = json.load(f)
    inc = re.compile(cfg.get("include") or r"(?!)")
    exc = re.compile(cfg["exclude"]) if cfg.get("exclude") else None
    senders = {nfc(x) for x in cfg.get("senders", [])}

    def match(name):
        if exc and exc.search(name):
            return False
        found = re.findall(r"_'([^']*)'", name)
        return bool(inc.search(name)) or (bool(found) and found[-1] in senders)

    return match


def parse_args(argv, default_out):
    """<eml 폴더> [출력 폴더] [--list 목록] [--rules 규칙.json] [--by month|year] [--hwp-pdf]"""
    argv = list(argv)
    opts = {}
    for flag in ("--by", "--list", "--rules"):
        if flag in argv:
            i = argv.index(flag)
            opts[flag] = argv[i + 1]
            del argv[i:i + 2]
    for flag in ("--hwp-pdf",):
        if flag in argv:
            opts[flag] = True
            argv.remove(flag)
    if not argv:
        raise SystemExit(__doc__ or "eml 폴더를 지정하세요.")
    src = Path(argv[0])
    out = Path(argv[1]) if len(argv) > 1 else Path(default_out)
    return src, out, opts


def select_files(src, opts):
    """폴더의 eml 가운데 --list / --rules 에 맞는 것만 고른다. (고른 파일, 목록에만 있는 이름)"""
    only = None
    if "--list" in opts:
        with open(opts["--list"], encoding="utf-8-sig") as f:
            only = {nfc(line.split("\t")[0].strip()) for line in f if line.strip()}
            only.discard("파일명")
    rule = load_rules(opts["--rules"]) if "--rules" in opts else None
    picked = []
    for p in sorted(src.rglob("*.eml")):
        name = nfc(p.name)
        if only is not None and name not in only:
            continue
        if rule is not None and not rule(name):
            continue
        picked.append(p)
    missing = sorted(only - {nfc(p.name) for p in picked}) if only is not None else []
    return picked, missing
