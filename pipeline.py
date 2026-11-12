"""
에이전트 파이프라인 — 웹(app.py)과 검증(evaluate.py)이 함께 쓰는 핵심 로직

  입력 파일
    ├─ 보도자료 → [에이전트 1: 기사 작성 + 판정] → 결과 정리(유형·원문·기사·판정)
    │                → 판정이 '승인'이면 [에이전트 2: 인터랙티브 기사]
    └─ 기사     →                     [에이전트 2: 인터랙티브 기사]

각 단계에서 무엇이 들어가고 무엇이 나왔는지 item["trace"]에 그대로 쌓아 화면에서 보여줌.
"""
import io
import json
import os
import re
import time
import zipfile
from xml.sax.saxutils import escape

from dotenv import load_dotenv
from openai import OpenAI
from hwp_reader import HwpError, hangul_to_text

load_dotenv()


def _env(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


# 에이전트 설정 — 2번 에이전트는 다른 조직/프로젝트일 수 있어 키·프로젝트를 따로 둘 수 있음
AGENTS = {
    "writer": {
        "label": "기사 작성 에이전트",
        "id": _env("UPSTAGE_AGENT_ID"),
        "config": _env("UPSTAGE_CONFIG_ID"),
        "key": _env("UPSTAGE_API_KEY"),
        "project": _env("UPSTAGE_PROJECT_ID"),
    },
    "interactive": {
        "label": "인터랙티브 기사 에이전트",
        "id": _env("UPSTAGE_AGENT2_ID"),
        "config": _env("UPSTAGE_AGENT2_CONFIG_ID"),
        "key": _env("UPSTAGE_AGENT2_API_KEY") or _env("UPSTAGE_API_KEY"),
        "project": _env("UPSTAGE_AGENT2_PROJECT_ID") or _env("UPSTAGE_PROJECT_ID"),
    },
}


def agent_ready(key: str) -> bool:
    a = AGENTS[key]
    return bool(a["id"] and a["key"])


ALLOWED_EXT = {"pdf", "hwp", "hwpx", "docx", "pptx", "xlsx", "png", "jpg", "jpeg", "bmp", "tiff", "heic", "txt"}

# Studio '보도자료에서 기사로' 워크플로우의 '보도자료유형분류' 노드와 같은 목록 (others → 기타)
PRESS_TYPES = [
    "사업·제도 시행", "성과·실적 홍보", "행사·일정 안내", "통계·조사 발표",
    "협약·MOU", "인사·동정", "해명·설명", "안전·재난 안내", "기타",
]
# 분류는 되지만 '이 유형으로 다시 만들기' 대상은 아닌 것 (예외처리 노드 → 인식불가/보도자료가 아님)
EXCEPTION_TYPE = "예외문서"
ALL_TYPES = PRESS_TYPES + [EXCEPTION_TYPE]

KIND_LABEL = {"press": "보도자료", "article": "기사"}


# ──────────────────────────────────────────────
# 입력 준비
# ──────────────────────────────────────────────
def text_to_docx(text: str) -> bytes:
    """텍스트를 최소한의 Word(.docx) 파일로 변환 (에이전트는 TXT를 받지 않음)."""
    paras = "".join(
        f'<w:p><w:r><w:t xml:space="preserve">{escape(line)}</w:t></w:r></w:p>'
        for line in text.splitlines()
    )
    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '</Types>'),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/></Relationships>'),
        "word/document.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{paras}</w:body></w:document>'),
    }
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in files.items():
            z.writestr(name, content)
    return buf.getvalue()


def _size(n: int) -> str:
    return f"{n / 1024 / 1024:.1f}MB" if n >= 1024 * 1024 else f"{max(1, round(n / 1024))}KB"


# ──────────────────────────────────────────────
# Upstage Studio 에이전트 호출
# ──────────────────────────────────────────────
def _client(a: dict) -> OpenAI:
    headers = {"x-project-id": a["project"]} if a["project"] else None
    return OpenAI(api_key=a["key"], base_url="https://api.upstage.ai/v2", default_headers=headers)


def _steps(job) -> list[tuple[str, str]]:
    """Job 결과의 output을 [(단계 이름, 텍스트), ...]로 정리."""
    steps = []
    for item in getattr(job, "output", None) or []:
        name = getattr(item, "model", None) or getattr(item, "type", None) or ""
        texts = [t for c in getattr(item, "content", None) or [] if (t := getattr(c, "text", None))]
        if texts:
            steps.append((str(name), "\n".join(texts)))
    return steps


def run_agent(key: str, filename: str, data: bytes, extra_text: str | None = None,
              preview: str | None = None) -> dict:
    """파일 업로드 → 에이전트 Job 실행 → 완료까지 대기.
    무엇을 보냈고 무엇을 받았는지 기록한 '단계(stage)' dict를 돌려줌. 실패해도 예외 대신 error에 기록."""
    a = AGENTS[key]
    stage = {
        "kind": "agent", "key": key, "label": a["label"],
        "agent_id": a["id"], "config_id": a["config"] or "최신 버전",
        "input_file": filename, "input_size": _size(len(data)), "input_text": extra_text,
        "input_preview": preview, "file_id": None, "job_id": None, "status": "",
        "seconds": 0, "steps": [], "error": None,
    }
    t0 = time.time()
    try:
        if not agent_ready(key):
            raise RuntimeError(f"{a['label']} 설정이 .env에 없습니다.")
        client = _client(a)
        uploaded = client.files.create(file=(filename, data), purpose="user_data")
        stage["file_id"] = uploaded.id

        content = [{"type": "input_file", "file_id": uploaded.id}]
        if extra_text:
            content.append({"type": "input_text", "text": extra_text})
        extra = {"config_id": a["config"]} if a["config"] else None
        include = ["all"]  # 중간 단계(Parse·Classify 등) 결과도 모두 받기
        job = client.responses.create(model=a["id"], include=include,
                                      input=[{"role": "user", "content": content}], extra_body=extra)
        stage["job_id"] = job.id
        for _ in range(150):  # 최대 약 5분 대기
            if job.status not in ("queued", "in_progress"):
                break
            time.sleep(2)
            job = client.responses.retrieve(job.id, include=include)
        stage["status"] = job.status
        if job.status != "completed":
            raise RuntimeError(f"에이전트 실행 실패 (status={job.status}, error={job.error})")
        stage["steps"] = [{"name": n, "text": t} for n, t in _steps(job)]
        if not stage["steps"]:
            raise RuntimeError("에이전트가 결과를 돌려주지 않았습니다. Studio 워크플로우를 확인하세요.")
    except Exception as e:
        stage["error"] = str(e)
    stage["seconds"] = round(time.time() - t0, 1)
    return stage


# ──────────────────────────────────────────────
# 에이전트 1 결과 해석 (원문·유형·기사)
# ──────────────────────────────────────────────
def _plain(text: str) -> str:
    """Parse 결과(JSON/HTML/마크다운)를 사람이 읽을 원문 텍스트로."""
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            content = data.get("content", data)
            if isinstance(content, dict):
                text = content.get("markdown") or content.get("text") or content.get("html") or text
            elif isinstance(content, str):
                text = content
        elif isinstance(data, str):
            text = data
    except (ValueError, json.JSONDecodeError):
        pass
    text = re.sub(r"<br\s*/?>", "\n", text)
    text = re.sub(r"</(p|h\d|li|tr|div)>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"^#+\s*", "", text, flags=re.M)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _norm(s: str) -> str:
    return re.sub(r"[\s·/\-_.,]", "", str(s)).lower()


def normalize_type(s: str) -> str | None:
    """'협약/MOU', '협약 MOU' 처럼 조금 다르게 적어도 표준 유형 이름으로 맞춤."""
    n = _norm(s)
    if not n:
        return None
    for t in ALL_TYPES:
        if _norm(t) == n:
            return t
    if n == "others":
        return "기타"
    for t in ALL_TYPES:
        if _norm(t) in n or n in _norm(t):
            return t
    return s.strip()


def detect_type(steps: list[dict]) -> tuple[str | None, str | None]:
    """최종 단계를 뺀 단계들에서 보도자료 유형 이름을 찾음. (유형, 찾은 단계 이름)"""
    keys = {_norm(t): t for t in ALL_TYPES}
    keys[_norm("others")] = "기타"
    middle = sorted(steps[:-1], key=lambda s: 0 if re.search(r"classif|분류", s["name"], re.I) else 1)
    for s in middle:
        text = s["text"]
        try:  # JSON 안의 한글이 \uXXXX로 이스케이프된 경우 풀어서 비교
            text = json.dumps(json.loads(text), ensure_ascii=False)
        except (ValueError, json.JSONDecodeError):
            pass
        for hay in (_norm(text[:500]), _norm(s["name"])):
            for k, v in keys.items():
                if k and k in hay:
                    return v, s["name"]
    return None, None


def find_source(steps: list[dict]) -> tuple[str, str | None]:
    """Parse(문서 인식) 단계의 결과를 원문으로 사용. (원문, 단계 이름)"""
    for s in steps:
        if re.search(r"parse|인식", s["name"], re.I):
            return _plain(s["text"]), s["name"]
    if len(steps) > 1:
        return _plain(steps[0]["text"]), steps[0]["name"]
    return "", None


def _clean(text: str) -> str:
    text = re.sub(r"【[^】]*】", "", str(text))  # 출처 표시 제거
    text = text.replace("**", "")
    return re.sub(r"[ \t]{2,}", " ", text).strip()


def _unwrap_json(raw):
    """JSON으로 여러 번 감싸져 온 경우 벗겨내기 (코드펜스 포함)."""
    for _ in range(3):
        if not isinstance(raw, str):
            break
        s = raw.strip()
        m = re.match(r"^```(?:json)?\s*(.*?)\s*```$", s, re.S)
        if m:
            s = m.group(1)
        try:
            raw = json.loads(s)
        except (ValueError, json.JSONDecodeError):
            return s
    return raw


def to_article(raw) -> dict:
    """에이전트 최종 출력(JSON 객체, JSON 문자열, 마크다운)을 기사 dict로."""
    raw = _unwrap_json(raw)
    if isinstance(raw, dict):
        if "headline" in raw or "body" in raw:
            body = raw.get("body", [])
            if isinstance(body, str):
                body = body.split("\n")
            titles = (raw.get("headlines") or raw.get("titles")
                      or raw.get("headline_candidates") or [])
            if isinstance(titles, str):
                titles = [titles]
            return {
                "headline": _clean(raw.get("headline", "")),
                "titles": [_clean(t) for t in titles if str(t).strip()],
                "lead": _clean(raw.get("lead", "")),
                "body": [_clean(p) for p in body if str(p).strip()],
            }
        texts = [v for v in raw.values() if isinstance(v, str)]
        raw = max(texts, key=len) if texts else json.dumps(raw, ensure_ascii=False)

    text = str(raw).replace("\\n", "\n")
    headline, body, titles = "", [], []
    for line in text.splitlines():
        line = line.strip()
        if not line or line == "---":
            continue
        m = re.match(r"^(?:제목\s*후보|후보|제목)\s*\d*\s*[:：.)]\s*(.+)$", line)
        if m:
            titles.append(_clean(m.group(1)))
            continue
        if line.startswith("# ") and not headline:
            headline = _clean(line[2:])
        else:
            cleaned = _clean(line.lstrip("#").lstrip("-").strip())
            if cleaned:
                body.append(cleaned)
    if not headline and body:
        headline = body.pop(0)
    return {"headline": headline or "생성된 기사", "titles": titles,
            "lead": body.pop(0) if len(body) > 1 else "", "body": body}


def title_candidates(article: dict) -> list[str]:
    """제목 후보 최대 3개 (에이전트가 후보를 주지 않으면 제목 1개)."""
    out = []
    for t in [article.get("headline", "")] + article.get("titles", []):
        if t and t not in out:
            out.append(t)
    return out[:3]


def article_text(item: dict) -> str:
    a = item["article"]
    title = item["titles"][0] if item.get("titles") else a["headline"]
    return "\n\n".join([title] + ([a["lead"]] if a["lead"] else []) + a["body"])


# ──────────────────────────────────────────────
# 판정 · 인터랙티브 결과 해석
#   에이전트 1 '보도자료에서 기사로' (v17)
#     Parse → 보도자료유형분류 → 유형별 추출 → 유형별 기사 작성 → 지시-1(재료 점수 → 승인/보류/거부)
#     (others·예외문서 경로에는 판정 노드가 없음 → 판정 없음)
#   서비스: 판정이 '승인'일 때만 에이전트 2 실행
#   에이전트 2 '인터랙티브 기사' (v3)
#     Parse → 가_원문정보_추출 → 나_기사이해 → 라_유형결정 → 분류-1(아래 5개)
#       → 지시-1/3/4/5(단일 HTML 웹페이지) 또는 지시-6(others: 보류 문구)
# ──────────────────────────────────────────────
DECISIONS = ["가능", "검토 필요", "불가"]
_DECISION_ALIASES = {
    "가능": "가능", "승인": "가능",
    "검토 필요": "검토 필요", "검토필요": "검토 필요", "검토": "검토 필요", "보류": "검토 필요",
    "불가": "불가", "거부": "불가",
}
# 분류-1 노드의 클래스 → 화면·정답지에 쓰는 짧은 이름
INTERACTIVE_TYPES = {
    "독자 참여형 시뮬레이터·계산기형 기사": "계산기형",
    "공공데이터 전수검증 대시보드형 기사": "대시보드형",
    "멀티미디어 스크롤리텔링(Scrollytelling) 기사": "타임라인형",
    "공간 기반 인터랙티브 맵·아카이브 기사": "지도형",
    "others": "해당없음",
}
# 정답지에 줄여 쓰거나 다르게 써도 알아보도록 (앞에서부터 확인)
_ITYPE_HINTS = [("시뮬레이터", "계산기형"), ("계산기", "계산기형"), ("대시보드", "대시보드형"),
                ("차트", "대시보드형"), ("전수검증", "대시보드형"), ("타임라인", "타임라인형"),
                ("스크롤리텔링", "타임라인형"), ("scrollytelling", "타임라인형"),
                ("지도", "지도형"), ("맵", "지도형"), ("아카이브", "지도형"),
                ("others", "해당없음"), ("해당없음", "해당없음"), ("기타", "해당없음"), ("없음", "해당없음")]


def _short(text: str) -> str:
    """노드 출력이 '"가능"', '{"value": "가능"}', '```가능```' 처럼 감싸져 와도 알맹이만."""
    data = _unwrap_json(text)
    if isinstance(data, dict) and len(data) == 1:
        data = next(iter(data.values()))
    return str(data).strip().strip("\"'`").strip()


def normalize_decision(s) -> str | None:
    t = _short(str(s or ""))
    if not t:
        return None
    for k, v in _DECISION_ALIASES.items():
        if t.startswith(k):
            return v
    if re.fullmatch(r"[0-5](\.0)?", t):  # 점수만 온 경우: 3~5 가능, 2 검토 필요, 0~1 불가
        n = int(float(t))
        return "가능" if n >= 3 else ("검토 필요" if n == 2 else "불가")
    return t


def normalize_itype(s) -> str | None:
    t = _short(str(s or ""))
    if not t:
        return None
    if t in INTERACTIVE_TYPES:
        return INTERACTIVE_TYPES[t]
    if t in INTERACTIVE_TYPES.values():
        return t
    low = t.lower()
    for hint, name in _ITYPE_HINTS:
        if hint in low:
            return name
    return t


def _is_decision(s: dict) -> bool:
    name = s.get("name", "").lower()
    if "validate" in name or "검증" in name:
        return True
    return len(s["text"]) <= 80 and normalize_decision(s["text"]) in DECISIONS


def detect_decision(steps: list[dict]) -> tuple[str | None, str | None]:
    """판정 노드 출력(가능/검토 필요/불가). 노드 이름과 상관없이 값으로 찾음. (판정, 단계 이름)"""
    # 1. Extract 점수와 Validate 체크 확인 (PDF 10~11페이지 기준)
    score_val = None
    score_from = None
    val_from = None
    val_can_pass = None

    for s in steps:
        txt = s.get("text", "")
        data = _unwrap_json(txt)
        if isinstance(data, dict):
            checks = data.get("checks")
            if isinstance(checks, list) and checks:
                val_from = s.get("name")
                for c in checks:
                    if "인터랙티브 가능" in str(c.get("name", "")):
                        val_can_pass = bool(c.get("passed"))
            sc = data.get("인터랙티브 제작 가능성 점수")
            if sc is None:
                sc = data.get("인터랙티브_점수")
            if sc is not None:
                try:
                    score_val = int(float(str(sc).strip()))
                    score_from = s.get("name")
                except (ValueError, TypeError):
                    pass

    if score_val is not None:
        dec = "가능" if score_val >= 3 else ("검토 필요" if score_val == 2 else "불가")
        return dec, val_from or score_from
    if val_can_pass is not None:
        return ("가능" if val_can_pass else "불가"), val_from

    # 2. 기존 텍스트 매칭
    for s in reversed(steps):
        if len(s["text"]) > 80:
            continue
        d = normalize_decision(s["text"])
        if d in DECISIONS:
            return d, s["name"]
    return None, None


def detect_itype(steps: list[dict]) -> tuple[str | None, str | None]:
    """분류-1 노드 출력(인터랙티브 유형). (유형, 단계 이름)"""
    for s in steps:
        t = _short(s["text"])
        if len(t) > 120:
            continue
        for full, short in INTERACTIVE_TYPES.items():
            if _norm(full) == _norm(t) or _norm(full) in _norm(t):
                return short, s["name"]
    return None, None


def find_html(steps: list[dict]) -> str | None:
    """에이전트 출력에서 웹페이지(HTML)를 찾음 → 화면에서 미리보기."""
    for s in reversed(steps):
        candidates = [s["text"]]
        data = _unwrap_json(s["text"])
        if isinstance(data, dict):
            candidates += [v for v in data.values() if isinstance(v, str)]
        for t in candidates:
            m = re.search(r"```html\s*(.*?)```", t, re.S | re.I)
            if m:
                return m.group(1).strip()
            m = re.search(r"(<!doctype html.*?</html>|<html.*?</html>)", t, re.S | re.I)
            if m:
                return m.group(1)
    return None


# ──────────────────────────────────────────────
# 전체 처리
# ──────────────────────────────────────────────

_META_FACT_KEYS = {
    "원문 맥락 서사", "원문_맥락_서사",
    "인터랙티브 제작 가능성 점수", "인터랙티브_제작_가능성_점수", "인터랙티브_점수",
}


def extract_facts(steps: list[dict]) -> dict:
    """워크플로우의 'Information Extract' (유형별 정보 추출) 단계에서 추출된 JSON 사실정보를 파싱."""
    facts = {}
    for s in steps:
        name = s.get("name", "") if isinstance(s, dict) else s[0]
        text = s.get("text", "") if isinstance(s, dict) else s[1]
        if any(k in name.lower() for k in ("validate", "검증")):
            continue
        parsed = _unwrap_json(text)
        if isinstance(parsed, dict) and (
            any(k in name.lower() for k in ("extract", "추출", "특징", "정보"))
            or any(k in parsed for k in ("인터랙티브 제작 가능성 점수", "원문 맥락 서사", "사업명", "성과명", "행사명", "지표명", "협약명", "인물목록", "자료구분", "재난유형"))
        ):
            # 기사 JSON(headline/body)은 제외
            if "headline" in parsed and "body" in parsed and "인터랙티브 제작 가능성 점수" not in parsed:
                continue
            facts.update(parsed)
    return facts


def extract_validation_results(steps: list[dict]) -> dict:
    """워크플로우의 'Validate' (자동 검증) 단계에서 평가/검증 결과를 파싱."""
    for s in steps:
        name = s.get("name", "") if isinstance(s, dict) else s[0]
        text = s.get("text", "") if isinstance(s, dict) else s[1]
        if any(k in name.lower() for k in ["validate", "검증"]):
            data = _unwrap_json(text)
            if isinstance(data, dict):
                return data
    return {}


def _nonempty(v) -> bool:
    return v is not None and str(v).strip() not in ("", "[]", "{}")


def _join_nonempty(items: list, sep: str = ", ") -> str:
    return sep.join(str(x).strip() for x in items if _nonempty(x))


def _parse_json_val(v):
    """문자열 내부에 직렬화된 JSON('{...}', '[...]')이 있으면 파이썬 객체로 복원."""
    for _ in range(2):
        if not isinstance(v, str):
            break
        s = v.strip()
        if not ((s.startswith("{") and s.endswith("}")) or (s.startswith("[") and s.endswith("]"))):
            break
        try:
            v = json.loads(s)
        except (ValueError, TypeError):
            break
    return v


def _format_dict_val(d: dict) -> str:
    """구조화된 dict 필드를 한국어 가독 문자열로 변환."""
    if "start" in d or "end" in d:
        return _join_nonempty([d.get("start"), d.get("end")], " ~ ")
    if "name" in d or "address" in d:
        n, a = _clean_fact_val(d.get("name")), _clean_fact_val(d.get("address"))
        return f"{n} ({a})" if (n and a) else (n or a)
    if "host" in d or "organizer" in d:
        parts = []
        if _nonempty(d.get("host")):
            parts.append(f"주최: {d['host']}")
        if _nonempty(d.get("organizer")):
            parts.append(f"주관: {d['organizer']}")
        return " / ".join(parts)
    if "dept" in d or "tel" in d:
        dep, tel = _clean_fact_val(d.get("dept")), _clean_fact_val(d.get("tel"))
        return f"{dep} ({tel})" if (dep and tel) else (dep or tel)
    if "time" in d or "title" in d:
        head = _join_nonempty([d.get("time"), d.get("title")], " · ")
        sp = _clean_fact_val(d.get("speaker"))
        return f"{head} ({sp})" if (head and sp) else head
    return ", ".join(f"{k}: {c}" for k, val in d.items() if (c := _clean_fact_val(val)))


def _clean_fact_val(v) -> str:
    """추출 스키마 값(JSON 문자열·dict·list 포함)을 사람이 읽을 한국어 텍스트로 정리."""
    v = _parse_json_val(v)
    if v is None:
        return ""
    if isinstance(v, list):
        return ", ".join(c for x in v if (c := _clean_fact_val(x)))
    if isinstance(v, dict):
        return _format_dict_val(v)
    s = str(v).strip()
    return "" if s.lower() in ("none", "null", "n/a", "해당없음", "없음", "-") else s


def _fact_get(facts: dict, *keys: str):
    """키의 띄어쓰기·밑줄·가운뎃점 변형까지 허용해 스키마 값을 조회."""
    for k in keys:
        if k in facts and _nonempty(facts[k]):
            return facts[k]
        for alt in (k.replace(" ", "_"), k.replace("·", "_"), k.replace(" ", "")):
            if alt in facts and _nonempty(facts[alt]):
                return facts[alt]
    return None


def _split_fact_items(v, split_comma: bool = False) -> list[str]:
    """배열(JSON 문자열 포함)이나 구분자 텍스트를 개별 항목 리스트로 분리."""
    parsed = _parse_json_val(v)
    if isinstance(parsed, list):
        return [c for x in parsed if (c := _clean_fact_val(x))]
    txt = _clean_fact_val(parsed)
    if not txt:
        return []
    sep_pat = r"[\n;·•,]" if split_comma else r"[\n;•]"
    return [part.strip(" -·•") for part in re.split(sep_pat, txt) if part.strip(" -·•")]


_TYPE_WIDGET_META = {
    "사업·제도 시행": ("policy", "내 지원 자격 자가진단"),
    "행사·일정 안내": ("event", "행사 장소 · 참여 안내"),
    "통계·조사 발표": ("stats_compare", "핵심 수치 비교 및 조사 해석"),
    "성과·실적 홍보": ("achievement", "주요 성과 및 선정 사유"),
    "협약·MOU": ("mou", "협약 당사자 및 협력 구조"),
    "해명·설명": ("clarification", "보도 쟁점 · 기관 입장 대조"),
    "안전·재난 안내": ("safety", "재난 안전 행동요령 체크리스트"),
    "인사·동정": ("brief", "주요 인물 · 행위 브리핑"),
    "기타": ("brief", "핵심 팩트 브리핑"),
}


def _build_stats(facts: dict, unique_nums: list[str], found_locs: list[str]) -> list[dict]:
    priority_stat_keys = [
        ("지원금액한도", "지원 한도"), ("총예산", "총 예산"), ("금액규모", "금액 규모"),
        ("지자체부담분", "지자체 부담분"), ("투자지원규모", "투자 규모"), ("핵심수치", "핵심 수치"),
        ("증감", "전년 대비 증감"), ("비교기준값", "비교 기준"), ("경쟁률/순위", "경쟁률·순위"),
        ("피해규모", "피해 규모"), ("위기경보단계", "위기경보"), ("참가규모", "참가 규모"),
        ("참가비", "참가비"), ("신청기간", "신청 기간"), ("행사일시", "행사 일시"),
        ("체결일시", "체결 일시"), ("달성일자", "달성 일자"), ("지원 대상", "지원 대상"),
        ("온오프라인구분", "진행 방식"), ("행사진행상태", "진행 상태"),
    ]
    stats, used_vals = [], set()
    for fk, label in priority_stat_keys:
        val = _clean_fact_val(_fact_get(facts, fk))
        if val and val not in used_vals and len(val) <= 45:
            stats.append({"label": label, "val": val})
            used_vals.add(val)
        if len(stats) >= 4:
            return stats
    for idx_n, num_val in enumerate(unique_nums):
        if len(stats) >= 4:
            break
        if num_val not in used_vals:
            stats.append({"label": f"핵심 수치 {idx_n + 1}", "val": num_val})
            used_vals.add(num_val)
    if len(stats) < 4 and found_locs:
        loc_summary = ", ".join(found_locs[:3])
        if loc_summary not in used_vals:
            stats.append({"label": "주요 지역", "val": loc_summary})
    return stats


def _build_timeline(facts: dict) -> list[dict]:
    raw_prog = _parse_json_val(_fact_get(facts, "주요프로그램"))
    if isinstance(raw_prog, list) and raw_prog:
        steps = []
        for idx, item in enumerate(raw_prog, 1):
            if isinstance(item, dict):
                t = _clean_fact_val(item.get("time") or item.get("구분")) or f"단계 {idx}"
                desc = _join_nonempty([item.get("title") or item.get("내용") or "", item.get("speaker") or item.get("비고") or ""], " · ")
                if desc:
                    steps.append({"title": t, "desc": desc})
            else:
                txt = _clean_fact_val(item)
                if txt:
                    steps.append({"title": f"프로그램 {idx}", "desc": txt})
        if steps:
            return steps

    timeline_candidates = [
        ("일정 · 접수", _fact_get(facts, "신청기간", "행사일시", "체결일시", "발생일시", "집계기간")),
        ("방법 · 절차", _fact_get(facts, "신청방법", "신청방법·기간", "선정방식", "대응조치")),
        ("혜택 · 결과", _fact_get(facts, "지원내용", "기대효과", "주요협력내용", "행동요령")),
        ("후속 · 계획", _fact_get(facts, "다음단계일정", "후속일정", "향후계획", "후속조치")),
    ]
    return [{"title": title, "desc": c} for title, desc in timeline_candidates if (c := _clean_fact_val(desc))]


def _build_type_widgets(detected: str | None, facts: dict) -> dict:
    """8대 보도자료 유형별 전용 인터랙티브 위젯 페이로드 생성."""
    widget_type, widget_title = _TYPE_WIDGET_META.get(detected or "기타", ("brief", "핵심 팩트 브리핑"))
    cond_policy = _clean_fact_val(_fact_get(facts, "지원자격", "지원 대상"))
    res_policy = _clean_fact_val(_fact_get(facts, "지원내용", "지원금액한도"))
    if widget_type == "brief" and cond_policy and res_policy:
        widget_type, widget_title = "policy", "내 지원 자격 자가진단"

    self_check = None
    cards: list[dict] = []
    compare_pairs: list[dict] = []
    checklist: dict | None = None

    def _add_card(tag: str, val, tone: str = "neutral"):
        txt = _clean_fact_val(val)
        if txt:
            cards.append({"tag": tag, "title": tag, "body": txt, "tone": tone})

    if widget_type == "policy":
        questions = []
        if cond_policy:
            questions.append({
                "id": "q_cond",
                "prompt": "지원 대상·자격 요건에 해당하시나요?",
                "yes_label": f"예 · {cond_policy[:60]}{'...' if len(cond_policy) > 60 else ''}",
                "no_label": "아니오 · 조건 미충족",
            })
        proc_val = _clean_fact_val(_fact_get(facts, "제출서류", "신청기간", "신청방법"))
        if proc_val:
            questions.append({
                "id": "q_proc",
                "prompt": "신청 기간·방법 및 제출서류를 확인하셨나요?",
                "yes_label": f"예 · {proc_val[:60]}{'...' if len(proc_val) > 60 else ''}",
                "no_label": "아니오 · 확인 필요",
            })
        if questions:
            self_check = {
                "title": widget_title,
                "questions": questions,
                "pass_heading": "지원 대상 요건 충족",
                "pass_detail": res_policy or "제시된 자격 요건과 신청 절차를 확인했습니다. 기간 내 신청을 진행하세요.",
                "fail_heading": "세부 요건 재확인이 필요합니다",
                "fail_detail": f"지원 자격({cond_policy or '본문 요건'}) 및 신청 기간·서류를 다시 확인해 주세요.",
                "highlight": _clean_fact_val(_fact_get(facts, "지원금액한도", "총예산")),
            }
        _add_card("달라진 점", _fact_get(facts, "달라진 점"), "accent")
        _add_card("재원 · 근거법령", _join_nonempty([_clean_fact_val(_fact_get(facts, "재원구성")), _clean_fact_val(_fact_get(facts, "근거법령"))], " / "))

    elif widget_type == "event":
        venues = _split_fact_items(_fact_get(facts, "행사장소"), split_comma=True)
        for idx, v in enumerate(venues, 1):
            cards.append({"tag": f"개최 장소 {idx}" if len(venues) > 1 else "개최 장소", "title": v, "body": _clean_fact_val(_fact_get(facts, "교통·주차안내")) or "세부 일정·참가 안내 참고", "tone": "accent"})
        _add_card("참가 대상 · 자격", _fact_get(facts, "참가대상·자격"))
        _add_card("참석 주요 인사", _fact_get(facts, "참석인사"))
        _add_card("주최 · 주관", _fact_get(facts, "주최·주관"))

    elif widget_type == "stats_compare":
        metric_name = _clean_fact_val(_fact_get(facts, "지표명", "조사명")) or "핵심 지표 비교"
        base_v = _clean_fact_val(_fact_get(facts, "비교기준값"))
        curr_v = _clean_fact_val(_fact_get(facts, "핵심수치"))
        delta_v = _clean_fact_val(_fact_get(facts, "증감"))
        if base_v or curr_v:
            compare_pairs.append({
                "badge": delta_v or "수치 비교",
                "title": metric_name,
                "left_label": "비교 기준값",
                "left_text": base_v or "원문 기준 참고",
                "right_label": "이번 조사 수치",
                "right_text": curr_v or _clean_fact_val(_fact_get(facts, "세부수치")) or "기사 본문 참고",
            })
        _add_card("세부 수치 · 추이", _join_nonempty([_clean_fact_val(_fact_get(facts, "세부수치")), _clean_fact_val(_fact_get(facts, "시계열추이"))], " / "), "accent")
        _add_card("기관 해석", _fact_get(facts, "기관해석"))
        _add_card("조사 범위 · 방법", _join_nonempty([_clean_fact_val(_fact_get(facts, "조사범위")), _clean_fact_val(_fact_get(facts, "조사방법"))], " · "))

    elif widget_type == "achievement":
        _add_card("성과명 · 선정 기관", _join_nonempty([_clean_fact_val(_fact_get(facts, "성과명")), _clean_fact_val(_fact_get(facts, "수여·선정기관"))], " — "), "accent")
        _add_card("핵심 선정 사유", _fact_get(facts, "핵심선정사유"))
        _add_card("경쟁률 · 연속 여부", _join_nonempty([_clean_fact_val(_fact_get(facts, "경쟁률/순위")), _clean_fact_val(_fact_get(facts, "연속여부"))], " / "))
        _add_card("기대 효과", _fact_get(facts, "기대효과"))

    elif widget_type == "mou":
        for party in _split_fact_items(_fact_get(facts, "협약당사자"), split_comma=True):
            cards.append({"tag": "협약 당사자", "title": party, "body": _clean_fact_val(_fact_get(facts, "기관별역할")) or "공동 협력 추진", "tone": "accent"})
        _add_card("주요 협력 내용", _fact_get(facts, "주요협력내용"))
        _add_card("이행 구속력", _fact_get(facts, "이행구속력"))

    elif widget_type == "clarification":
        compare_pairs.append({
            "badge": _clean_fact_val(_fact_get(facts, "자료구분")) or "팩트 대조",
            "title": _clean_fact_val(_fact_get(facts, "쟁점")) or "보도 내용과 기관 입장 비교",
            "left_label": f"보도 요지 ({_clean_fact_val(_fact_get(facts, '대상보도매체')) or '대상 보도'})",
            "left_text": _clean_fact_val(_fact_get(facts, "보도요지")) or _clean_fact_val(_fact_get(facts, "대상보도제목")) or "보도 내용 참고",
            "right_label": "기관 공식 입장 · 반박",
            "right_text": _clean_fact_val(_fact_get(facts, "기관입장")) or _clean_fact_val(_fact_get(facts, "부인사항")) or "기관 해명 참고",
        })
        if _clean_fact_val(_fact_get(facts, "인정사항")) or _clean_fact_val(_fact_get(facts, "부인사항")):
            compare_pairs.append({
                "badge": "인정 vs 부인",
                "title": "사실 확인 구분",
                "left_label": "인정한 사항",
                "left_text": _clean_fact_val(_fact_get(facts, "인정사항")) or "명시 없음",
                "right_label": "부인 · 정정 사항",
                "right_text": _clean_fact_val(_fact_get(facts, "부인사항")) or "명시 없음",
            })
        _add_card("수치 대비 · 근거자료", _join_nonempty([_clean_fact_val(_fact_get(facts, "수치대비")), _clean_fact_val(_fact_get(facts, "근거자료"))], " / "), "accent")

    elif widget_type == "safety":
        action_items = _split_fact_items(_fact_get(facts, "행동요령"), split_comma=True)
        measure_items = _split_fact_items(_fact_get(facts, "대응조치"), split_comma=True)
        control_items = _split_fact_items(_fact_get(facts, "통제제한사항"), split_comma=True)
        chk_items = action_items or measure_items
        if chk_items or control_items:
            checklist = {
                "title": f"{_clean_fact_val(_fact_get(facts, '재난유형')) or '재난·안전'} {'행동요령' if action_items else '대응·안전'} 점검",
                "items": chk_items,
                "warn_items": control_items,
            }
        areas = _split_fact_items(_fact_get(facts, "발생장소"), split_comma=True)
        vuln_txt = _clean_fact_val(_fact_get(facts, "취약대상"))
        for area in areas:
            cards.append({"tag": "영향·관리 지역", "title": area, "body": vuln_txt or "안전 유의 필요", "tone": "accent"})
        if not areas and vuln_txt:
            _add_card("중점 보호 · 취약대상", vuln_txt, "accent")
        if action_items and measure_items:
            _add_card("기관 대응 조치 · 투입 자원", _join_nonempty([_clean_fact_val(_fact_get(facts, "대응조치")), _clean_fact_val(_fact_get(facts, "투입자원"))], " / "))
        elif _clean_fact_val(_fact_get(facts, "투입자원")):
            _add_card("투입 자원 · 지원 규모", _fact_get(facts, "투입자원"))
        _add_card("관련 통계 · 향후 계획", _join_nonempty([_clean_fact_val(_fact_get(facts, "관련통계")), _clean_fact_val(_fact_get(facts, "향후계획"))], " / "))

    else:
        _add_card("주요 인물 · 행위", _join_nonempty([_clean_fact_val(_fact_get(facts, "인물목록")), _clean_fact_val(_fact_get(facts, "행위유형"))], " — "), "accent")
        _add_card("논의 안건 · 주요 발언", _join_nonempty([_clean_fact_val(_fact_get(facts, "논의안건")), _clean_fact_val(_fact_get(facts, "주요발언(인용)"))], " / "))

    return {
        "widget_type": widget_type,
        "widget_title": widget_title,
        "self_check": self_check,
        "cards": cards,
        "compare_pairs": compare_pairs,
        "checklist": checklist,
    }


def _enrich_from_article_text(
    article_text: str,
    widgets: dict,
    facts: dict,
    unique_nums: list[str],
    found_locs: list[str],
    timeline: list[dict],
) -> None:
    """others 경로(facts={}) 또는 스키마 빈약 문서에서도 본문 기반 비교·카드·체크리스트를 보강."""
    sentences = [s.strip() for s in re.split(r"(?<=[.다요])\s+|\n+", article_text) if len(s.strip()) > 16 and not s.strip().startswith("#")]

    # 1) 본문 내 '기존 X에서 Y로' 전후 비교 패턴 자동 추출
    if not widgets["compare_pairs"]:
        m_cmp = re.search(
            r"(?:기존|당초|지난해|전년)\s*([0-9][0-9,.]*\s*(?:조\s*원|억\s*원|만\s*원|원|%|명|개소|곳|건))에서\s*([0-9][0-9,.]*\s*(?:조\s*원|억\s*원|만\s*원|원|%|명|개소|곳|건))",
            article_text,
        )
        if m_cmp:
            b_val, a_val = m_cmp.group(1).strip(), m_cmp.group(2).strip()
            ctx = next((s for s in sentences if m_cmp.group(1) in s and m_cmp.group(2) in s), "지원·정책 규모 변경 비교")
            widgets["compare_pairs"].append({
                "badge": f"{b_val} → {a_val}",
                "title": "전·후 지원/수치 변화 비교",
                "left_label": "변경 전 (기존)",
                "left_text": b_val,
                "right_label": "변경 후 (확대·개편)",
                "right_text": f"{a_val} ({ctx[:80]}{'...' if len(ctx) > 80 else ''})",
            })

    # 2) safety 체크리스트 보강 (추출값이 비어있으면 본문 내 점검·예방·지원 문장으로 체크리스트 구성)
    if widgets["widget_type"] == "safety" and not widgets["checklist"]:
        safety_sents = [s for s in sentences if any(k in s for k in ("점검", "운영", "지원", "예방", "대피", "수칙", "주의", "유의", "감시"))]
        if safety_sents:
            widgets["checklist"] = {
                "title": "안전 관리 · 예방 대책 확인 체크리스트",
                "items": safety_sents[:4],
                "warn_items": [],
            }

    # 3) 카드가 비어 있을 때(특히 others 경로) 본문 핵심 팩트 카드 자동 생성
    if not widgets["cards"]:
        if found_locs:
            widgets["cards"].append({
                "tag": "관련 지역",
                "title": ", ".join(found_locs[:5]),
                "body": next((s for s in sentences if any(loc in s for loc in found_locs)), "기사 본문에 언급된 주요 대상 지역입니다."),
                "tone": "accent",
            })
        sup_sent = next((s for s in sentences if any(k in s for k in ("지원", "투입", "확대", "운영", "점검", "추진"))), "")
        if sup_sent:
            widgets["cards"].append({
                "tag": "핵심 조치 · 지원",
                "title": (unique_nums[0] + " 등 핵심 추진" if unique_nums else "주요 추진 내용"),
                "body": sup_sent,
                "tone": "neutral",
            })

    # 4) 타임라인도 비어 있으면 본문 기승전결 핵심 문장 3개 이내로 요약 타임라인 보강
    if not timeline and len(sentences) >= 2:
        labels = ["배경 · 현황", "주요 대책 · 실행", "지원 · 기대효과"]
        for lbl, sent in zip(labels, sentences[:3]):
            timeline.append({"title": lbl, "desc": sent})


def _num_to_float(num_str: str) -> float | None:
    m = re.match(r"([0-9][0-9,.]*)\s*(.*)", num_str.strip())
    if not m:
        return None
    val = float(m.group(1).replace(",", ""))
    unit = re.sub(r"\s+", "", m.group(2))
    mult = {"억원": 1e8, "만원": 1e4, "조원": 1e12}.get(unit, 1.0)
    return val * mult


def _build_visual_bars(compare_pairs: list[dict], unique_nums: list[str], article_text: str) -> list[dict]:
    """수치 전후 비교 또는 주요 정량 지표를 애니메이션 비교 막대 데이터로 변환."""
    bars = []
    # 1) '기존 X -> Y' 직접 수치 비교가 있으면 비율 막대 생성
    m_cmp = re.search(
        r"(?:기존|당초|지난해|전년)\s*([0-9][0-9,.]*\s*(?:조\s*원|억\s*원|만\s*원|원|%|명|개소|곳|건))에서\s*([0-9][0-9,.]*\s*(?:조\s*원|억\s*원|만\s*원|원|%|명|개소|곳|건))",
        article_text,
    )
    if m_cmp:
        a_s, b_s = m_cmp.group(1).strip(), m_cmp.group(2).strip()
        av, bv = _num_to_float(a_s), _num_to_float(b_s)
        if av and bv and max(av, bv) > 0:
            mx = max(av, bv)
            delta_pct = round(((bv - av) / av) * 100)
            sign = "+" if delta_pct >= 0 else ""
            bars.append({
                "title": "전·후 수치 변화 비교",
                "badge": f"{sign}{delta_pct}%",
                "items": [
                    {"label": "변경 전", "val": a_s, "pct": max(14, round(av / mx * 100)), "accent": False},
                    {"label": "변경 후", "val": b_s, "pct": max(14, round(bv / mx * 100)), "accent": True},
                ],
            })
            return bars

    # 2) 단일 비교 스키마(통계·조사 발표 등)에서 숫자 2개가 잡히면 막대 생성
    for cp in compare_pairs:
        nums_l = re.findall(r"[0-9][0-9,.]*\s*(?:조\s*원|억\s*원|만\s*원|원|%|명|개소|곳|건)", cp.get("left_text", ""))
        nums_r = re.findall(r"[0-9][0-9,.]*\s*(?:조\s*원|억\s*원|만\s*원|원|%|명|개소|곳|건)", cp.get("right_text", ""))
        if nums_l and nums_r:
            av, bv = _num_to_float(nums_l[0]), _num_to_float(nums_r[0])
            if av and bv and max(av, bv) > 0:
                mx = max(av, bv)
                bars.append({
                    "title": cp["title"],
                    "badge": cp["badge"],
                    "items": [
                        {"label": cp["left_label"], "val": nums_l[0], "pct": max(14, round(av / mx * 100)), "accent": False},
                        {"label": cp["right_label"], "val": nums_r[0], "pct": max(14, round(bv / mx * 100)), "accent": True},
                    ],
                })
                return bars

    # 3) 일반 다수 수치 기사는 주요 수치 규모 랭킹 막대로 시각화
    if len(unique_nums) >= 3:
        sents = [s.strip() for s in re.split(r"(?<=[.다요])\s+|\n+", article_text) if s.strip()]
        picked = []
        for n in unique_nums[:4]:
            ctx = next((s for s in sents if n in s and not s.startswith("#")), "")
            short_lbl = re.sub(r"\s+", " ", ctx.split(n)[0])[-16:].strip(" ·,-") if ctx else ""
            picked.append((short_lbl or n, n))
        bar_items = [
            {"label": lbl, "val": val, "pct": max(28, 100 - i * 18), "accent": i == 0}
            for i, (lbl, val) in enumerate(picked)
        ]
        bars.append({"title": "기사 속 주요 정량 지표", "badge": f"{len(bar_items)}개 지표", "items": bar_items})
    return bars


def _build_explore_tabs(article_text: str) -> list[dict]:
    """기사 본문 문단들을 주제별 클릭 탐색 탭(내 관심 항목 골라보기)으로 그룹핑."""
    paras = [p.strip() for p in article_text.split("\n\n") if len(p.strip()) > 25 and not p.strip().startswith("#")]
    if len(paras) < 2:
        return []
    themes = [
        ("시설 · 점검", ("시설", "점검", "아케이드", "하우스", "도로", "장소", "현장")),
        ("지원 · 예산", ("지원", "예산", "만원", "억", "기금", "요금", "난방", "혜택", "한도")),
        ("대상 · 취약계층", ("대상", "노인", "취약", "이재민", "돌봄", "쉼터", "참가", "시민")),
        ("안전 · 예방·일정", ("예방", "화재", "단속", "수칙", "신청", "일정", "기간", "안내")),
    ]
    tabs, used = [], set()
    for label, kws in themes:
        matched = [p for i, p in enumerate(paras) if i not in used and any(k in p for k in kws)]
        if matched:
            for p in matched:
                used.add(paras.index(p))
            nums = re.findall(r"\d[\d,.]*\s*(?:억원|만원|%|명|개소|곳|건)", " ".join(matched))
            tabs.append({
                "id": f"tab_{len(tabs)}",
                "label": label,
                "highlight": nums[0] if nums else "",
                "paragraphs": matched[:2],
            })
    if len(tabs) < 2:
        tabs = [
            {"id": f"tab_{i}", "label": f"핵심 포인트 {i + 1}", "highlight": "", "paragraphs": [p]}
            for i, p in enumerate(paras[:3])
        ]
    return tabs[:4]


def _build_reader_quiz(article_text: str, unique_nums: list[str]) -> dict | None:
    """기사 본문의 실제 수치 1개를 바탕으로 독자 참여형 30초 팩트체크 퀴즈 생성."""
    if not unique_nums:
        return None
    sents = [s.strip() for s in re.split(r"(?<=[.다요])\s+|\n+", article_text) if 20 <= len(s.strip()) <= 180 and not s.strip().startswith("#")]
    target_num, target_sent = None, None
    for n in unique_nums:
        sent = next((s for s in sents if n in s), None)
        if sent:
            target_num, target_sent = n, sent
            break
    if not target_num or not target_sent:
        return None

    m_unit = re.match(r"([0-9][0-9,.]*)\s*(.*)", target_num)
    if not m_unit:
        return None
    base_val = float(m_unit.group(1).replace(",", ""))
    unit = m_unit.group(2).strip()
    is_int = "." not in m_unit.group(1)

    def _fmt(v: float) -> str:
        s = f"{int(round(v)):,}" if is_int else f"{round(v, 1):,}"
        return f"{s}{unit}" if unit in ("%", "배", "위", "명", "건", "회", "곳", "개소", "개", "대", "톤", "만원", "억원") else f"{s} {unit}".strip()

    options = [target_num]
    # 같은 기사의 다른 수치 중 단위가 같거나 다른 숫자를 먼저 활용하고, 부족하면 배율 보완
    for cand in unique_nums:
        if cand != target_num and cand not in options and cand.endswith(unit):
            options.append(cand)
        if len(options) >= 4:
            break
    for mult in (0.5, 2.0, 1.5, 0.25):
        if len(options) >= 4:
            break
        cand = _fmt(max(1.0, base_val * mult))
        if cand not in options:
            options.append(cand)
    if len(options) < 3:
        return None

    # 결정적 순서 배치 (새로고침해도 동일 보기 유지)
    shift = sum(ord(ch) for ch in target_num) % len(options)
    ordered = options[shift:] + options[:shift]
    ans_idx = ordered.index(target_num)
    masked = target_sent.replace(target_num, "[  ?  ]", 1)
    return {
        "question": f"기사 내용 중 [ ? ]에 들어갈 알맞은 수치는 무엇일까요?",
        "prompt": masked,
        "options": ordered[:4],
        "answer_idx": ans_idx,
        "answer": target_num,
        "explanation": target_sent,
    }


def _extract_calendar_event(article_headline: str, facts: dict) -> dict | None:
    """행사·신청 기간에서 YYYY-MM-DD 날짜가 확인되면 .ics 캘린더 저장용 메타 생성."""
    raw_dt = _clean_fact_val(_fact_get(facts, "행사일시", "신청기간", "체결일시"))
    if not raw_dt:
        return None
    dates = re.findall(r"(20\d{2})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})", raw_dt)
    if not dates:
        return None
    y, m, d = int(dates[0][0]), int(dates[0][1]), int(dates[0][2])
    start_iso = f"{y:04d}{m:02d}{d:02d}"
    if len(dates) >= 2:
        y2, m2, d2 = int(dates[1][0]), int(dates[1][1]), int(dates[1][2])
        end_iso = f"{y2:04d}{m2:02d}{d2:02d}"
    else:
        end_iso = start_iso
    return {
        "title": _clean_fact_val(_fact_get(facts, "행사명", "사업명", "협약명")) or article_headline,
        "start": start_iso,
        "end": end_iso,
        "location": _clean_fact_val(_fact_get(facts, "행사장소", "신청방법", "체결장소")),
        "label": raw_dt,
    }


def _strip_dates_and_phones(text: str) -> str:
    """PDF 10페이지: 수치 기준에서 날짜·전화번호 제외."""
    cleaned = re.sub(r"\d{2,4}[-.]\d{3,4}[-.]\d{4}", " ", text)  # 전화번호
    cleaned = re.sub(r"\d{4}년(?:\s*\d{1,2}월)?(?:\s*\d{1,2}일)?", " ", cleaned)  # 연·월·일
    cleaned = re.sub(r"\d{1,2}월\s*\d{1,2}일", " ", cleaned)  # 월·일
    return cleaned


def evaluate_interactive_criteria(
    article_text: str,
    facts: dict,
    val_data: dict | None = None,
    detected: str | None = None,
) -> dict:
    """
    발표 설계도(PDF 10~11페이지) 기준 인터랙티브 판정 및 4단계 자동 검증:
      - 5대 기준 (0~5점): 수치(3개↑, 날짜/전화 제외) · 비교 · 장소(2곳↑, 청사 제외) · 시간(3단계↑) · 맞춤(조건+결과)
      - 점수 구간: 3~5점 가능 (4~5점 GREEN, 3점 경고), 2점 검토 필요, 0~1점 불가
      - 자동 검증 4단계: 기사 생성 · 유형 확인 · 인터랙티브 가능(≥3) · GREEN 조건(≥4)
    """
    val_checks = (val_data or {}).get("checks", [])

    # 예외문서('보도자료 아님' 안내)는 기사/인터랙티브 대상에서 즉시 제외
    if detected == EXCEPTION_TYPE:
        auto_checks = [
            {"name": "기사 생성", "passed": False, "level": "error", "detail": "예외문서로 분류되어 기사 생성 제외"},
            {"name": "유형 확인", "passed": False, "level": "error", "detail": "보도자료 아님 (예외문서 안내 출력)"},
            {"name": "인터랙티브 가능", "passed": False, "level": "fail", "detail": "예외문서 — 대상 제외"},
            {"name": "GREEN 조건", "passed": False, "level": "warning", "detail": "해당 없음"},
        ]
        return {
            "score": 0,
            "decision": "불가",
            "decision_label": "보도자료 아님 (예외문서)",
            "grade": "EXCEPTION",
            "is_eligible": False,
            "is_review": False,
            "itype": "해당없음",
            "recommended_formats": [],
            "criteria": [],
            "auto_checks": auto_checks,
            "stats": [],
            "timeline": [],
            "fact_items": [],
            "cta": {},
            "facts": {},
            "val_checks": val_checks,
            "widget_type": "exception",
            "widget_title": "보도자료 아님 (예외문서)",
            "self_check": None,
            "cards": [],
            "compare_pairs": [],
            "checklist": None,
            "explore_tabs": [],
            "visual_bars": [],
            "reader_quiz": None,
            "calendar_event": None,
            "raw_html": None,
            "skipped": "예외문서('보도자료 아님' 안내)로 분류되어 인터랙티브 대상에서 제외되었습니다.",
        }

    # 1. Extract 노드의 공식 점수 확인
    agent_score = facts.get("인터랙티브 제작 가능성 점수")
    if agent_score is None:
        agent_score = facts.get("인터랙티브_점수")
    parsed_agent_score = None
    if agent_score is not None:
        try:
            parsed_agent_score = max(0, min(5, int(float(str(agent_score).strip()))))
        except (ValueError, TypeError):
            parsed_agent_score = None

    # Validate 노드 체크값 확인 (GREEN / yellow / 인터랙티브 가능)
    val_can_pass = any("인터랙티브 가능" in str(c.get("name", "")) and c.get("passed") for c in val_checks)
    val_green_pass = any(
        any(k in str(c.get("name", "")).lower() for k in ("green", "yellow")) and c.get("passed")
        for c in val_checks
    )

    # 2. 5대 기준 계산 (PDF 10페이지 기준)
    # ① 수치: 3개 이상 (금액·인원·비율 등 의미가 다른 숫자, 날짜·전화번호 제외)
    num_source = _strip_dates_and_phones(article_text)
    num_matches = re.findall(
        r"(\d[\d,.]*\s*(?:조\s*원|억\s*원|만\s*원|원|%|퍼센트|명|인|개월|개사|개소|개|곳|건|회|배|위|톤|㎢|㎡|km|m|시간|일|년))",
        num_source,
    )
    unique_nums = []
    seen_num_norm = set()
    for n in num_matches:
        clean = re.sub(r"\s+", "", n)
        if clean not in seen_num_norm:
            seen_num_norm.add(clean)
            unique_nums.append(n.strip())
    fact_num_keys = ("지원금액한도", "총예산", "금액규모", "지자체부담분", "경쟁률/순위", "핵심수치", "참가규모", "투자지원규모", "피해규모", "기부금액·물품")
    has_fact_nums = sum(1 for k in fact_num_keys if _clean_fact_val(facts.get(k)))
    crit_num = {
        "name": "수치 (3개↑)",
        "passed": len(unique_nums) >= 3 or has_fact_nums >= 2,
        "detail": (
            "수치 검출(날짜·전화 제외): " + ", ".join(unique_nums[:4])
            if unique_nums else "금액·인원·비율 등 수치 데이터 부족"
        ),
    }

    # ② 비교: 두 값 이상 비교 (전년 대비·지역별·제도 변경 전후·증감률·순위)
    cmp_terms = ["전년", "대비", "증가", "감소", "인상", "인하", "확대", "축소", "비교", "전보다", "증감", "순위", "달라진", "기존", "개편"]
    found_cmp = [k for k in cmp_terms if k in article_text]
    has_fact_cmp = any(_clean_fact_val(facts.get(k)) for k in ("달라진 점", "비교기준값", "증감", "수치대비", "이전협약이력"))
    crit_cmp = {
        "name": "비교 (전년·지역·전후)",
        "passed": bool(found_cmp) or has_fact_cmp,
        "detail": f"비교 지표 확인 ('{found_cmp[0]}')" if found_cmp else ("추출 스키마 내 비교/변동 항목 확인" if has_fact_cmp else "전년 대비·제도 전후 등 비교 지표 부재"),
    }

    # ③ 장소: 2곳 이상 (지도에 찍을 수 있는 지명·시설·도로, 기관 청사 제외)
    loc_terms = [
        "대구", "경북", "포항", "경주", "안동", "구미", "김천", "영천", "상주", "문경",
        "경산", "의성", "청송", "영양", "영덕", "청도", "고령", "성주", "칠곡", "예천",
        "봉화", "울진", "울릉", "서울", "부산", "울산", "광주", "대전", "인천", "수원",
        "청주", "전주", "춘천", "제주", "창원", "김해", "진주", "세종",
    ]
    found_locs = [loc for loc in loc_terms if loc in article_text]
    fact_loc = _clean_fact_val(_fact_get(facts, "행사장소", "체결장소", "발생장소", "장소"))
    loc_places = _split_fact_items(_fact_get(facts, "행사장소", "발생장소"), split_comma=True)
    crit_loc = {
        "name": "장소 (2곳↑)",
        "passed": len(found_locs) >= 2 or len(loc_places) >= 2 or (len(found_locs) >= 1 and bool(fact_loc)),
        "detail": f"지명 {len(found_locs)}곳 확인: " + ", ".join(found_locs[:5]) if found_locs else ("장소 정보 확인: " + fact_loc if fact_loc else "지도 표기 가능한 복수 지명·시설 부재"),
    }

    # ④ 시간: 3단계 이상 (일정표·절차 신청→심사→지급, 시점별 추이)
    time_terms = ["공고", "모집", "신청", "접수", "심사", "평가", "선정", "발표", "지급", "시행", "종료", "추진"]
    found_times = [k for k in time_terms if k in article_text]
    dates = re.findall(r"\d{1,2}월\s*\d{1,2}일", article_text)
    raw_prog = _parse_json_val(_fact_get(facts, "주요프로그램"))
    prog_count = len(raw_prog) if isinstance(raw_prog, list) else 0
    has_fact_time = sum(1 for k in ("신청기간", "선정방식", "다음단계일정", "후속일정", "향후계획", "시계열추이", "행사일시", "협약기간") if _clean_fact_val(_fact_get(facts, k)))
    crit_time = {
        "name": "시간 (3단계↑)",
        "passed": len(found_times) >= 3 or len(dates) >= 2 or has_fact_time >= 2 or prog_count >= 3,
        "detail": "절차·추이 확인: " + " → ".join(found_times[:4]) if found_times else ("일정·단계 정보 확인" if has_fact_time else "3단계 이상 일정·시계열 추이 부재"),
    }

    # ⑤ 맞춤: 조건 + 결과 (나이·소득·거주지 등 조건에 따라 받을 금액·할 일이 명시)
    cond_val = _clean_fact_val(facts.get("지원자격") or facts.get("지원 대상") or facts.get("참가대상·자격") or facts.get("취약대상"))
    res_val = _clean_fact_val(facts.get("지원내용") or facts.get("지원금액한도") or facts.get("행동요령") or facts.get("참가비") or facts.get("주요프로그램"))
    has_condition = any(k in article_text for k in ["대상", "자격", "기준", "조건", "세 이상", "세 이하", "이내", "거주", "소득", "취약"]) or bool(cond_val)
    has_result = any(k in article_text for k in ["지원", "지급", "혜택", "선정", "제공", "면제", "감면", "할인", "행동요령"]) or bool(res_val)
    crit_custom = {
        "name": "맞춤 (조건+결과)",
        "passed": has_condition and has_result,
        "detail": "자격·조건과 수혜·행동 결과 연계 확인" if (has_condition and has_result) else "독자 맞춤형 조건·결과 불명확",
    }

    criteria = [crit_num, crit_cmp, crit_loc, crit_time, crit_custom]
    calculated_score = sum(1 for c in criteria if c["passed"])

    # 최종 점수 확정 (1차 워크플로우 Extract 점수 우선)
    final_score = parsed_agent_score if parsed_agent_score is not None else calculated_score

    # 3. PDF 10페이지 판정 규칙: 3~5점 가능 / 2점 검토 필요 / 0~1점 불가
    if final_score >= 3:
        decision = "가능"
        is_eligible = True
        is_review = False
        if final_score >= 4 or val_green_pass:
            grade = "GREEN"
            decision_label = "가능 (GREEN · 자동 추천)"
        else:
            grade = "WARNING"
            decision_label = "가능 (경고 · 4점 미만)"
    elif final_score == 2:
        decision = "검토 필요"
        is_eligible = False
        is_review = True
        grade = "REVIEW"
        decision_label = "검토 필요 (기자 검토 후 제작 여부 결정)"
    else:
        decision = "불가"
        is_eligible = False
        is_review = False
        grade = "FAIL"
        decision_label = "불가 (일반 텍스트 기사로 작성)"

    # 4. PDF 11페이지 자동 검증 4단계 (기사 생성 · 유형 확인 · 인터랙티브 가능 · GREEN 조건)
    article_ok = bool(article_text and len(article_text.strip()) > 30)
    type_ok = bool(detected and detected != EXCEPTION_TYPE)
    ia_ok = val_can_pass if (val_checks and parsed_agent_score is None) else (final_score >= 3)
    green_ok = val_green_pass if (val_checks and parsed_agent_score is None) else (final_score >= 4)
    auto_checks = [
        {
            "name": "기사 생성",
            "passed": article_ok,
            "level": "pass" if article_ok else "error",
            "detail": "유형별 프롬프트 기사 초안 작성 완료" if article_ok else "기사 초안 생성 실패",
        },
        {
            "name": "유형 확인",
            "passed": type_ok,
            "level": "pass" if type_ok else "error",
            "detail": f"분류 유형 확정 ({detected})" if type_ok else "보도자료 유형 미확인",
        },
        {
            "name": "인터랙티브 가능",
            "passed": ia_ok,
            "level": "pass" if ia_ok else "fail",
            "detail": f"{final_score}점 ≥ 3점 (대상 포함)" if ia_ok else f"{final_score}점 < 3점 (인터랙티브 대상 제외)",
        },
        {
            "name": "GREEN 조건",
            "passed": green_ok,
            "level": "pass" if green_ok else "warning",
            "detail": f"{final_score}점 ≥ 4점 (GREEN)" if green_ok else f"{final_score}점 < 4점 (경고 · warning)",
        },
    ]

    # 5. 유형별 인터랙티브 위젯 & 추천 형식 결정
    widgets = _build_type_widgets(detected, facts)
    recommended_formats = []
    if widgets["widget_type"] == "policy":
        recommended_formats.append("자가진단 계산기")
    elif widgets["widget_type"] == "safety":
        recommended_formats.append("행동요령 체크리스트")
    elif widgets["widget_type"] == "clarification":
        recommended_formats.append("쟁점·입장 대조표")
    elif widgets["widget_type"] == "mou":
        recommended_formats.append("기관·협력 구조 카드")
    if (crit_num["passed"] and crit_cmp["passed"]) or detected in ("통계·조사 발표", "성과·실적 홍보"):
        recommended_formats.append("핵심 수치 차트·대시보드")
    if crit_loc["passed"] or detected in ("행사·일정 안내", "안전·재난 안내"):
        recommended_formats.append("지도·장소 안내")
    if crit_time["passed"] or detected == "행사·일정 안내":
        recommended_formats.append("단계별 타임라인")
    if not recommended_formats and is_eligible:
        recommended_formats.append("핵심 팩트 카드")

    if "자가진단 계산기" in recommended_formats:
        itype = "계산기형"
    elif "행동요령 체크리스트" in recommended_formats:
        itype = "체크리스트형"
    elif "쟁점·입장 대조표" in recommended_formats:
        itype = "비교대조형"
    elif "핵심 수치 차트·대시보드" in recommended_formats:
        itype = "대시보드형"
    elif "지도·장소 안내" in recommended_formats:
        itype = "지도형"
    elif "단계별 타임라인" in recommended_formats:
        itype = "타임라인형"
    else:
        itype = "해당없음"

    # 6. 정제된 팩트/통계/타임라인/체험형 모듈 조립
    fact_items = [
        {"key": k, "val": c}
        for k, v in facts.items()
        if k not in _META_FACT_KEYS and (c := _clean_fact_val(v))
    ]
    stats = _build_stats(facts, unique_nums, found_locs)
    timeline = _build_timeline(facts)
    _enrich_from_article_text(article_text, widgets, facts, unique_nums, found_locs, timeline)
    visual_bars = _build_visual_bars(widgets["compare_pairs"], unique_nums, article_text)
    explore_tabs = _build_explore_tabs(article_text)
    reader_quiz = _build_reader_quiz(article_text, unique_nums)
    calendar_event = _extract_calendar_event(article_text.splitlines()[0].lstrip("# "), facts)
    cta = {
        "contact": _clean_fact_val(_fact_get(facts, "문의처", "신고문의처", "담당부서")),
        "method": _clean_fact_val(_fact_get(facts, "신청방법", "신청방법·기간")),
        "link": _clean_fact_val(_fact_get(facts, "신청링크")),
    }

    return {
        "score": final_score,
        "decision": decision,
        "decision_label": decision_label,
        "grade": grade,
        "is_eligible": is_eligible,
        "is_review": is_review,
        "itype": itype,
        "recommended_formats": recommended_formats,
        "criteria": criteria,
        "auto_checks": auto_checks,
        "stats": stats,
        "timeline": timeline,
        "fact_items": fact_items,
        "cta": cta,
        "facts": facts,
        "val_checks": val_checks,
        "visual_bars": visual_bars,
        "explore_tabs": explore_tabs,
        "reader_quiz": reader_quiz,
        "calendar_event": calendar_event,
        "raw_html": None,
        **widgets,
    }


def _writer(item: dict, press_type: str | None):
    extra = None
    if press_type:
        extra = f"이 보도자료의 유형은 '{press_type}'입니다. 이 유형으로 분류해 그 형식에 맞는 기사를 작성하세요."
    stage = run_agent("writer", *item["file"], extra_text=extra)
    item["trace"].append(stage)
    if stage["error"]:
        raise RuntimeError(stage["error"])

    steps = stage["steps"]
    # 판정 노드가 마지막에 붙으므로, 기사는 판정이 아닌 마지막 출력
    art_step = next((x for x in reversed(steps) if not _is_decision(x)), steps[-1])
    article = to_article(art_step["text"])

    detected, type_from = detect_type(steps)
    effective_detected = detected or "기타"
    _, d_from = detect_decision(steps)

    # 팩트 및 Validate 검증 결과 추출
    facts = extract_facts(steps)
    val_data = extract_validation_results(steps)
    article_full_text = (
        f"# {article['headline']}\n\n"
        + (f"{article['lead']}\n\n" if article["lead"] else "")
        + "\n\n".join(article["body"])
    )
    interactive = evaluate_interactive_criteria(
        article_full_text, facts, val_data, detected=press_type or effective_detected
    )
    llm_html = find_html(steps)
    if llm_html:
        interactive["html"] = llm_html

    decision = interactive["decision"]
    src_from = "붙여넣은/TXT 원문 그대로"
    if not item["source_fixed"]:
        item["source"], src_from = find_source(steps)

    is_exception = (effective_detected == EXCEPTION_TYPE)
    exception_msg = (_short(art_step["text"])[:300] or "보도자료 형식이 아니거나 내용 판독이 어려운 예외문서입니다.") if is_exception else None

    item.update(
        article=article,
        titles=title_candidates(article),
        detected=effective_detected,
        type=press_type or effective_detected,
        has_lead=bool(article["lead"]),
        exception=exception_msg,
        decision=decision,
        facts=facts,
        val_data=val_data,
        interactive=interactive,
        llm_html=llm_html,
    )

    if effective_detected == "기타":
        type_desc = f"기타 (others → 정보 추출 건너뛰고 바로 기사 작성 · '{type_from}' 단계)" if type_from else "기타 (분류 결과를 못 찾아 '기타'로 처리)"
    elif is_exception:
        type_desc = f"예외문서 (보도자료 아님 안내 · '{type_from}' 단계)"
    else:
        type_desc = f"{item['detected']}" + (f"  ← '{type_from}' 단계에서 찾음" if type_from else "")

    item["trace"].append({
        "kind": "parse",
        "label": "결과 정리 (서비스가 에이전트 출력을 해석)",
        "rows": [
            ("보도자료 유형", type_desc),
            *([("사용자가 지정한 유형", press_type)] if press_type else []),
            ("원문", f"{len(item['source']):,}자  ← {src_from or '찾지 못함'}"),
            ("기사", f"'{art_step['name']}' 단계 출력 → 제목 {len(item['titles'])}개 · "
                   f"리드 {'있음' if article['lead'] else '없음'} · 본문 {len(article['body'])}문단"),
            ("추출 팩트", f"{len(interactive.get('fact_items', []))}개 스키마 항목 추출" if facts else "없음 (others/예외문서 경로)"),
            ("인터랙티브 판정", f"{interactive['decision_label']} ({interactive['score']}/5점)"
                                + (f"  ← '{d_from}' 단계" if d_from else "  ← 5대 기준 자동 산출")),
        ],
    })


def process(filename: str, data: bytes, kind: str = "press", press_type: str | None = None,
            with_interactive: bool = True, display_name: str | None = None) -> dict:
    """입력 파일 하나를 끝까지 처리 (단일 워크플로우: 파싱→분류→추출→작성→자동 검증)."""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext not in ALLOWED_EXT:
        raise ValueError(f"지원하지 않는 파일 형식입니다: .{ext}")

    item = {"kind": kind, "filename": display_name or filename, "orig": (filename, data),
            "source": "", "source_fixed": False, "trace": [], "article": None,
            "titles": [], "detected": None, "type": None, "interactive": None, "decision": None}
    if ext in ("hwp", "hwpx"):
        try:
            text = hangul_to_text(data)
            item.update(file=(filename.rsplit(".", 1)[0] + ".docx", text_to_docx(text)),
                        source=text, source_fixed=True)
        except Exception:
            item["file"] = (filename, data)
    elif ext == "txt":
        text = data.decode("utf-8-sig", errors="ignore")
        item.update(file=(filename.rsplit(".", 1)[0] + ".docx", text_to_docx(text)),
                    source=text, source_fixed=True)
    else:
        item["file"] = (filename, data)

    item["trace"].append({
        "kind": "input", "label": f"입력 — {KIND_LABEL[kind]}",
        "rows": [
            ("파일", display_name or filename),
            ("형식 · 크기", f"{ext.upper()} · {_size(len(data))}"),
            *([("변환", f"{ext.upper()} 문서를 {item['file'][0]}(Word)로 변환하여 전달")] if ext in ("txt", "hwp", "hwpx") else []),
            ("처리 경로", "단일 워크플로우 (파싱 → 10개 유형 분류 → 스키마 추출 → 기사 작성 → 5대 기준 자동 검증)"),
        ],
        "preview": item["source"] or None,
    })

    _writer(item, press_type)
    if with_interactive:
        interactive = item.get("interactive") or {}
        auto_summary = " · ".join(
            f"{c['name']}({'PASS' if c['passed'] else ('WARN' if c['level'] == 'warning' else 'FAIL')})"
            for c in interactive.get("auto_checks", [])
        )
        item["trace"].append({
            "kind": "gate",
            "label": "자동 검증 & 인터랙티브 판정 게이트 (5대 기준 · 0~5점)",
            "rows": [
                ("인터랙티브 점수", f"{interactive.get('score', 0)} / 5점"),
                ("최종 판정", interactive.get("decision_label", interactive.get("decision", "불가"))),
                ("추천 형식", ", ".join(interactive.get("recommended_formats") or []) or "일반 텍스트 기사"),
                ("자동 검증 (4단계)", auto_summary or "-"),
                ("충족 기준", ", ".join([c["name"] for c in interactive.get("criteria", []) if c.get("passed")]) or "없음"),
            ]
        })
        if interactive.get("is_eligible"):
            item["trace"].append({
                "kind": "parse",
                "label": "인터랙티브 템플릿 바인딩 (독자 참여형 위젯)",
                "rows": [
                    ("렌더링 방식", "추출 스키마(JSON) → 독립형 반응형 템플릿 실시간 바인딩"),
                    ("생성 요소", f"요약 지표 {len(interactive.get('stats', []))}개 · 타임라인 {len(interactive.get('timeline', []))}단계"
                                  + (" · 맞춤형 체크카드" if interactive.get("self_check") else "")),
                    ("외부 배포", "Standalone 단독 실행 HTML 소스코드 복사 · iframe 임베드 · 다운로드 지원"),
                ]
            })
    return item
