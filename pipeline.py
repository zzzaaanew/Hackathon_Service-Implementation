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
DECISIONS = ["승인", "보류", "거부"]
# 분류-1 노드의 클래스 → 화면·정답지에 쓰는 짧은 이름
INTERACTIVE_TYPES = {
    "독자 참여형 시뮬레이터·계산기형 기사": "계산기형",
    "공공데이터 전수검증 대시보드형 기사": "대시보드형",
    "멀티미디어 스크롤리텔링(Scrollytelling) 기사": "스크롤리텔링형",
    "공간 기반 인터랙티브 맵·아카이브 기사": "지도형",
    "others": "해당없음",
}
# 정답지에 줄여 쓰거나 다르게 써도 알아보도록 (앞에서부터 확인)
_ITYPE_HINTS = [("시뮬레이터", "계산기형"), ("계산기", "계산기형"), ("대시보드", "대시보드형"),
                ("전수검증", "대시보드형"), ("스크롤리텔링", "스크롤리텔링형"), ("scrollytelling", "스크롤리텔링형"),
                ("지도", "지도형"), ("맵", "지도형"), ("아카이브", "지도형"),
                ("others", "해당없음"), ("해당없음", "해당없음"), ("기타", "해당없음"), ("없음", "해당없음")]


def _short(text: str) -> str:
    """노드 출력이 '"승인"', '{"value": "승인"}', '```승인```' 처럼 감싸져 와도 알맹이만."""
    data = _unwrap_json(text)
    if isinstance(data, dict) and len(data) == 1:
        data = next(iter(data.values()))
    return str(data).strip().strip("\"'`").strip()


def normalize_decision(s) -> str | None:
    t = _short(str(s or ""))
    if not t:
        return None
    for d in DECISIONS:
        if t.startswith(d):
            return d
    if t.startswith("불가"):  # 예전 워크플로우·정답지 표기
        return "거부"
    if re.fullmatch(r"[0-5](\.0)?", t):  # enum 없이 점수만 온 경우: 4~5 승인, 3 보류, 이하 거부
        n = int(float(t))
        return "승인" if n >= 4 else ("보류" if n == 3 else "거부")
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
    """판정 노드 출력(승인/보류/거부). 노드 이름과 상관없이 값으로 찾음. (판정, 단계 이름)"""
    # 1. Validate 노드 또는 Extract 노드의 '인터랙티브 제작 가능성 점수' 확인
    for s in steps:
        txt = s.get("text", "")
        try:
            data = json.loads(txt)
            if isinstance(data, dict):
                checks = data.get("checks", [])
                if checks:
                    val_yellow = any("yellow" in c.get("name", "") and c.get("passed") for c in checks)
                    val_pass = any("인터랙티브 가능" in c.get("name", "") and c.get("passed") for c in checks)
                    if val_yellow:
                        return "승인", s["name"]
                    elif val_pass:
                        return "보류", s["name"]
                    else:
                        return "거부", s["name"]
                score = data.get("인터랙티브 제작 가능성 점수") or data.get("인터랙티브_점수")
                if score is not None:
                    n = int(float(str(score)))
                    dec = "승인" if n >= 4 else ("보류" if n == 3 else "거부")
                    return dec, s["name"]
        except Exception:
            pass

    # 2. 기존 텍스트 매칭
    for s in reversed(steps):
        if len(s["text"]) > 80:
            continue
        d = normalize_decision(s["text"])
        if d in DECISIONS:
            return d, s["name"]
    return None, None


def detect_itype(steps: list[dict]) -> tuple[str | None, str | None]:
    """분류-1 노드 출력(인터랙티브 유형). 승인일 때만 나옴. (유형, 단계 이름)
    나·라 노드의 긴 분석문에도 유형 이름이 나오므로 짧은 출력만 봄."""
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

def extract_facts(steps: list[dict]) -> dict:
    """1차 워크플로우의 'Information Extract' 단계에서 추출된 JSON 사실정보를 파싱."""
    facts = {}
    for s in steps:
        name = s.get("name", "") if isinstance(s, dict) else s[0]
        text = s.get("text", "") if isinstance(s, dict) else s[1]
        if "extract" in name.lower() or "특징" in name or "정보" in name:
            t = text.strip()
            if t.startswith("```"):
                t = t.strip("`").removeprefix("json").strip()
            try:
                parsed = json.loads(t)
                if isinstance(parsed, dict):
                    facts.update(parsed)
            except Exception:
                pass
    return facts


def extract_validation_results(steps: list[dict]) -> dict:
    """1차 워크플로우의 'Validate' 단계에서 평가/검증 결과를 파싱."""
    for s in steps:
        name = s.get("name", "") if isinstance(s, dict) else s[0]
        text = s.get("text", "") if isinstance(s, dict) else s[1]
        if any(k in name.lower() for k in ["validate", "검증"]):
            t = text.strip()
            try:
                data = json.loads(t)
                if isinstance(data, dict):
                    return data
            except Exception:
                pass
    return {}


def evaluate_interactive_criteria(article_text: str, facts: dict, val_data: dict | None = None) -> dict:
    """
    1차 워크플로우(Extract 점수 및 Validate 검증)와 5대 핵심 인터랙티브 지표 검증 합성
    """
    agent_score = facts.get("인터랙티브 제작 가능성 점수") or facts.get("인터랙티브_점수")
    parsed_agent_score = None
    if agent_score is not None:
        try:
            parsed_agent_score = int(float(str(agent_score).strip()))
        except Exception:
            parsed_agent_score = None

    val_checks = (val_data or {}).get("checks", [])
    val_can_pass = any(c.get("name") == "인터랙티브 가능" and c.get("passed") for c in val_checks)
    val_yellow_pass = any("yellow" in c.get("name", "") and c.get("passed") for c in val_checks)

    # 1. 수치 (3개 이상)
    num_matches = re.findall(r'(\d+[\d,.]*\s*(?:억\s*원|만\s*원|원|%|세|명|개사|개|곳|건|회))', article_text)
    unique_nums = []
    for n in num_matches:
        clean = re.sub(r'\s+', '', n)
        if clean not in [re.sub(r'\s+', '', u) for u in unique_nums]:
            unique_nums.append(n)
    crit_num = {
        "name": "수치 (3개↑)",
        "passed": len(unique_nums) >= 3,
        "detail": f"수치 {len(unique_nums)}개 검출: " + ", ".join(unique_nums[:5]) if unique_nums else "수치 데이터 부족"
    }

    # 2. 비교 (전년·지역·전후)
    cmp_terms = ["전년", "대비", "증가", "감소", "인상", "인하", "확대", "축소", "비교", "전보다", "대비해"]
    found_cmp = [k for k in cmp_terms if k in article_text]
    crit_cmp = {
        "name": "비교 (전년·지역)",
        "passed": len(found_cmp) > 0,
        "detail": f"비교 지표 발견 ('{found_cmp[0]}')" if found_cmp else "전년 대비 등 비교 지표 부재"
    }

    # 3. 장소 (2곳 이상)
    loc_terms = [
        "대구", "경북", "포항", "경주", "안동", "구미", "김천", "영천", "상주", "문경",
        "경산", "의성", "청송", "영양", "영덕", "청도", "고령", "성주", "칠곡", "예천",
        "봉화", "울진", "울릉", "서울", "부산"
    ]
    found_locs = list(set([loc for loc in loc_terms if loc in article_text]))
    crit_loc = {
        "name": "장소 (2곳↑)",
        "passed": len(found_locs) >= 2,
        "detail": f"지명 {len(found_locs)}곳 발견: " + ", ".join(found_locs) if found_locs else "지명 정보 부재"
    }

    # 4. 시간 (3단계 이상)
    time_terms = ["모집", "신청", "접수", "심사", "선정", "발표", "지급", "시행", "종료"]
    found_times = [k for k in time_terms if k in article_text]
    dates = re.findall(r'\d{1,2}월\s*\d{1,2}일', article_text)
    crit_time = {
        "name": "시간 (3단계↑)",
        "passed": len(found_times) >= 3 or len(dates) >= 2 or bool(facts.get("선정방식")),
        "detail": "추진 절차: " + " → ".join(found_times[:4]) if found_times else "일정 단계 부재"
    }

    # 5. 맞춤 (조건 + 결과)
    has_condition = any(k in article_text for k in ["대상", "자격", "기준", "조건", "세 이상", "세 이하", "이내", "거주"]) or bool(facts.get("지원자격"))
    has_result = any(k in article_text for k in ["지원", "지급", "혜택", "선정", "제공", "면제", "감면", "원"]) or bool(facts.get("지원내용"))
    crit_custom = {
        "name": "맞춤 (조건+결과)",
        "passed": has_condition and has_result,
        "detail": "조건(자격/대상)과 수혜 결과(지원금/서비스) 연계 확인" if (has_condition and has_result) else "맞춤 조건/결과 불명확"
    }

    criteria = [crit_num, crit_cmp, crit_loc, crit_time, crit_custom]
    calculated_score = sum(1 for c in criteria if c["passed"])

    if parsed_agent_score is not None:
        final_score = parsed_agent_score
    else:
        final_score = calculated_score

    if final_score >= 4 or val_yellow_pass:
        decision = "승인"
    elif final_score >= 3 or val_can_pass:
        decision = "보류"
    else:
        decision = "불가"

    is_eligible = (final_score >= 3)
    itype = "계산기형" if is_eligible else "해당없음"

    stats = []
    if facts.get("지원금액한도"):
        stats.append({"label": "지원 한도", "val": facts.get("지원금액한도")})
    elif unique_nums:
        stats.append({"label": "핵심 수치", "val": unique_nums[0]})

    if facts.get("총예산"):
        stats.append({"label": "총 사업 예산", "val": facts.get("총예산")})
    elif len(unique_nums) > 1:
        stats.append({"label": "총 규모", "val": unique_nums[1]})

    if facts.get("신청기간"):
        stats.append({"label": "신청 기간", "val": facts.get("신청기간").split("~")[-1].strip()})
    elif dates:
        stats.append({"label": "주요 일정", "val": dates[0]})

    target_val = facts.get("지원 대상") or ("도내 청년 창업기업" if "청년" in article_text else "도민 및 유관기업")
    stats.append({"label": "지원 대상", "val": target_val})

    timeline = [
        {"title": "1단계: 참여기업 모집 및 접수", "desc": facts.get("신청방법") or facts.get("신청기간", "K-Startup 누리집 온라인 신청 접수")},
        {"title": "2단계: 자격 검증 및 심사·평가", "desc": facts.get("선정방식", "요건 적합성 심사 및 사업화 계획 대면 발표 평가")},
        {"title": "3단계: 최종 선정 및 자금 교부", "desc": facts.get("지원내용", "기업당 사업화 자금 지급 및 1:1 맞춤형 멘토링 매칭")},
    ]

    return {
        "score": final_score,
        "decision": decision,
        "is_eligible": is_eligible,
        "itype": itype,
        "criteria": criteria,
        "stats": stats,
        "timeline": timeline,
        "facts": facts,
        "val_checks": val_checks,
        "self_check": is_eligible,
        "raw_html": None
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
    
    # 팩트 및 Validate 검증 결과 추출
    facts = extract_facts(steps)
    val_data = extract_validation_results(steps)
    article_full_text = f"# {article['headline']}\n\n" + (f"{article['lead']}\n\n" if article['lead'] else "") + "\n\n".join(article['body'])
    interactive = evaluate_interactive_criteria(article_full_text, facts, val_data)
    llm_html = find_html(steps)
    if llm_html:
        interactive["html"] = llm_html

    decision = interactive["decision"]
    src_from = "붙여넣은/TXT 원문 그대로"
    if not item["source_fixed"]:
        item["source"], src_from = find_source(steps)
    detected, type_from = detect_type(steps)
    item.update(article=article, titles=title_candidates(article),
                detected=detected or "기타", type=press_type or detected or "기타",
                has_lead=bool(article["lead"]),
                exception=_short(art_step["text"])[:200] if detected == EXCEPTION_TYPE else None,
                decision=decision,
                facts=facts,
                val_data=val_data,
                interactive=interactive,
                llm_html=llm_html)
    item["trace"].append({
        "kind": "parse", "label": "결과 정리 (서비스가 에이전트 출력을 해석)",
        "rows": [
            ("보도자료 유형", f"{item['detected']}" + (f"  ← '{type_from}' 단계에서 찾음" if type_from
                                                    else "  ← 분류 결과를 못 찾아 '기타'로 처리")),
            *([("사용자가 지정한 유형", press_type)] if press_type else []),
            ("원문", f"{len(item['source']):,}자  ← {src_from or '찾지 못함'}"),
            ("기사", f"'{art_step['name']}' 단계 출력 → 제목 {len(item['titles'])}개 · "
                   f"리드 {'있음' if article['lead'] else '없음'} · 본문 {len(article['body'])}문단"),
            ("인터랙티브 판정", f"{decision}  ← '{d_from}' 단계" if decision
             else "없음 (others·예외문서 경로이거나, 승인/보류/거부 출력을 못 찾음)"),
        ],
    })


# 보도자료를 넣었을 때 2번 에이전트에 무엇을 줄지: article(1번이 쓴 기사) / source(원본 보도자료)
AGENT2_INPUT = _env("UPSTAGE_AGENT2_INPUT", "article").lower()


def _interactive(item: dict):
    if item["kind"] == "press" and AGENT2_INPUT == "source":
        stage = run_agent("interactive", *item["file"], preview=item["source"] or None)
    elif item["kind"] == "press":
        text = article_text(item)
        file = ("article.docx", text_to_docx(text))
        stage = run_agent("interactive", *file, preview=text)
    else:
        stage = run_agent("interactive", *item["file"],
                          preview=item["source"] or None)
    item["trace"].append(stage)
    if stage["error"]:
        item["interactive"] = {"error": stage["error"]}
        return
    steps = stage["steps"]
    itype, t_from = detect_itype(steps)
    html = find_html(steps)
    item["interactive"] = {"itype": itype, "html": html, "raw": steps[-1]["text"]}
    item["trace"].append({
        "kind": "parse", "label": "인터랙티브 결과 정리 (서비스가 에이전트 출력을 해석)",
        "rows": [
            ("인터랙티브 유형", f"{itype}  ← '{t_from}' 단계" if itype else "찾지 못함"),
            ("웹페이지", f"HTML {len(html):,}자 → 아래에서 미리보기" if html
             else f"없음 — 마지막 '{steps[-1]['name']}' 출력은 일반 텍스트"),
        ],
    })


def process(filename: str, data: bytes, kind: str = "press", press_type: str | None = None,
            with_interactive: bool = True, display_name: str | None = None) -> dict:
    """입력 파일 하나를 끝까지 처리. 에이전트 1이 실패하면 예외, 에이전트 2 실패는 item에 기록."""
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
            *([("변환", f"TXT는 에이전트가 못 읽어서 {item['file'][0]}(Word)로 바꿔 보냄")] if ext == "txt" else []),
            ("처리 경로", ("기사 작성 에이전트(판정 포함)" + (" → 승인이면 인터랙티브 기사 에이전트"
                                                          if with_interactive else ""))
                       if kind == "press" else "인터랙티브 기사 에이전트 (판정 없이 바로)"),
        ],
        "preview": item["source"] or None,
    })

    if kind == "press":
        _writer(item, press_type)
        if with_interactive:
            interactive = item.get("interactive") or {}
            item["trace"].append({
                "kind": "gate",
                "label": "인터랙티브 적합도 판정 게이트 (5대 기준 종합 검증)",
                "rows": [
                    ("적합도 점수", f"{interactive.get('score', 0)} / 5점"),
                    ("최종 판정", interactive.get("decision", "불가")),
                    ("적격 여부", "인터랙티브 기사 적격 (3점 이상)" if interactive.get("is_eligible") else "인터랙티브 부적격 · 표준 기사 권장 (2점 이하)"),
                    ("충족 기준", ", ".join([c["name"] for c in interactive.get("criteria", []) if c.get("passed")]) or "없음"),
                ]
            })
            if interactive.get("is_eligible"):
                item["trace"].append({
                    "kind": "parse",
                    "label": "인터랙티브 컴포넌트 렌더링 (자가진단 계산기 · 통계 카드 · 타임라인)",
                    "rows": [
                        ("렌더링 방식", "클라이언트 반응형 컴포넌트 템플릿 실시간 바인딩"),
                        ("생성 위젯", "실시간 자가진단 계산기, 핵심 통계 지표 그리드, 3단계 일정 타임라인"),
                        ("외부 배포", "Standalone 단독 실행 HTML 소스코드 복사 및 다운로드 지원"),
                    ]
                })
    return item
