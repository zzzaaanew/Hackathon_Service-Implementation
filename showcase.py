"""
showcase.py — 시연용 보도자료 5건 전용 인터랙티브 기사(쇼케이스) 연결

일반 보도자료는 pipeline.py가 만든 추출 팩트를 templates/interactive.html에 바인딩한다.
아래 5건은 데모데이 시연용으로, 원문 수치를 사람이 직접 대조해 만든 전용 페이지
(templates/showcase/<slug>.html)를 대신 보여 준다. 기사 본문은 그대로 워크플로우 출력이다.

판별: 원문(또는 기사)에 고유 문구가 모두 들어 있으면 해당 건으로 본다.
      원문을 못 읽은 경우에만 파일명(setNN)으로 판별한다.
"""
import math
import re

SHOWCASES = {
    "housing": {
        "label": "맞춤 계산기형",
        "title": "월세 20만 원, 보증료 40만 원 — 나도 받을 수 있을까",
        "source": "경상북도 청년정책과 보도자료 (2026. 8. 9.)",
        "keys": ["청년월세", "전세보증금반환보증", "보증료"],
        "file": "set02",
    },
    "carfire": {
        "label": "스크롤 데이터 스토리형",
        "title": "여름, 차가 불탄다 — 경북 차량화재 5년의 기록",
        "source": "경상북도 소방본부 보도자료 (2026. 8. 25.)",
        "keys": ["차량화재", "537건", "과열"],
        "file": "set04",
    },
    "walk": {
        "label": "지도 탐색형",
        "title": "18곳이 손들었고, 5곳이 남았다 — 걷기 좋은 거리 지도",
        "source": "행정안전부 보도자료 (2026. 11. 9.)",
        "keys": ["보행환경종합정비", "18개지구"],
        "file": "set16",
    },
    "daegyeong": {
        "label": "독자 맞춤 탐색형",
        "title": "대구와 경북, 일자리로 잇는다 — 대경권 광역이음",
        "source": "고용노동부 보도자료 (2026. 11. 5.)",
        "keys": ["광역이음", "미래모빌리티", "대경권"],
        "file": "set12",
    },
    "medschool": {
        "label": "타임라인 · Q&A 탐색형",
        "title": "의대 없는 지역에 100개의 자리 — 60 대 40은 어떻게 정해졌나",
        "source": "교육부 · 보건복지부 보도자료 (2026. 11. 4.)",
        "keys": ["지역의대", "국립순천대", "국립경국대"],
        "file": "set11",
    },
}


def _squash(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def match(item: dict) -> str | None:
    """item(원문·기사·파일명)을 보고 쇼케이스 slug를 돌려줌. 해당 없으면 None."""
    article = item.get("article") or {}
    text = _squash(item.get("source", ""))
    if len(text) < 200:  # 원문 추출 실패 → 기사 본문으로 보완
        text += _squash(" ".join([article.get("headline", ""), article.get("lead", "")]
                                 + list(article.get("body", []))))
    for slug, sc in SHOWCASES.items():
        if all(k in text for k in sc["keys"]):
            return slug
    name = (item.get("orig") or ("",))[0] or item.get("filename", "")
    for slug, sc in SHOWCASES.items():
        if sc["file"] in name:
            return slug
    return None


# ──────────────────────────────────────────────
# 간이 한국 지도 (외부 지도 라이브러리 없이 SVG로 그림 — 단독 HTML 파일에서도 동작)
# ──────────────────────────────────────────────
_COAST = [  # (경도, 위도) 본토 해안선 근사 — 위치 안내용 약도
    (126.68, 37.95), (126.52, 37.75), (126.60, 37.45), (126.78, 37.20), (126.88, 37.00),
    (126.50, 36.98), (126.15, 36.78), (126.33, 36.45), (126.52, 36.20), (126.62, 36.00),
    (126.50, 35.75), (126.45, 35.50), (126.38, 35.20), (126.30, 34.95), (126.35, 34.75),
    (126.22, 34.50), (126.55, 34.30), (126.92, 34.42), (127.22, 34.50), (127.45, 34.52),
    (127.75, 34.70), (127.92, 34.90), (128.20, 34.85), (128.45, 34.80), (128.70, 34.88),
    (128.95, 35.05), (129.20, 35.15), (129.40, 35.45), (129.50, 35.80), (129.58, 36.05),
    (129.40, 36.12), (129.45, 36.60), (129.45, 37.05), (129.30, 37.35), (129.05, 37.70),
    (128.70, 38.10), (128.37, 38.60), (127.80, 38.30), (127.20, 38.30), (126.95, 38.10),
]
_K = 100.0                       # 위도 1도 = 100px
_KX = _K * math.cos(math.radians(36.2))
MAP_W, MAP_H = 330, 580


def xy(lon: float, lat: float) -> tuple[float, float]:
    return round((lon - 125.75) * _KX, 1), round((38.85 - lat) * _K, 1)


def korea_path() -> str:
    pts = [xy(*p) for p in _COAST]
    n = len(pts)
    d = f"M{(pts[0][0] + pts[1][0]) / 2:.1f},{(pts[0][1] + pts[1][1]) / 2:.1f}"
    for i in range(1, n + 1):  # 꼭짓점을 조절점으로 쓰는 2차 곡선 → 부드러운 윤곽
        p, q = pts[i % n], pts[(i + 1) % n]
        d += f" Q{p[0]},{p[1]} {(p[0] + q[0]) / 2:.1f},{(p[1] + q[1]) / 2:.1f}"
    return d + "Z"


PLACES = {
    "대구 수성구": (128.63, 35.84), "경기 파주시": (126.78, 37.76), "충북 음성군": (127.69, 36.94),
    "충남 아산시": (127.00, 36.79), "전북 김제시": (126.88, 35.80),
    "대구": (128.60, 35.87), "안동": (128.73, 36.57), "순천": (127.49, 34.95),
}


def context(slug: str) -> dict:
    """전용 템플릿에 넘길 공통 값."""
    jx, jy = xy(126.55, 33.38)
    return {
        "sc": SHOWCASES[slug], "slug": slug,
        "kr": {"w": MAP_W, "h": MAP_H, "path": korea_path(), "jeju": {"x": jx, "y": jy},
               "pt": {k: dict(zip("xy", xy(*v))) for k, v in PLACES.items()}},
    }
