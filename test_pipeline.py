"""Self-check verifying pipeline & template rendering against 성영하자_최종.pdf specs."""
import json
import sys
import uuid

sys.stdout.reconfigure(encoding="utf-8")

import app as web_app
import pipeline


def check_score_tiers():
    base_text = "경상북도 포항시와 구미시에서 지원 사업을 시행한다."
    # 5점 / 4점 -> 가능 (GREEN)
    r4 = pipeline.evaluate_interactive_criteria(base_text, {"인터랙티브 제작 가능성 점수": 4}, detected="사업·제도 시행")
    assert r4["decision"] == "가능" and r4["grade"] == "GREEN" and r4["is_eligible"], r4

    # 3점 -> 가능 (WARNING: 4점 미만 경고)
    r3 = pipeline.evaluate_interactive_criteria(base_text, {"인터랙티브 제작 가능성 점수": 3}, detected="행사·일정 안내")
    assert r3["decision"] == "가능" and r3["grade"] == "WARNING" and r3["is_eligible"], r3
    green_chk = next(c for c in r3["auto_checks"] if c["name"] == "GREEN 조건")
    ia_chk = next(c for c in r3["auto_checks"] if c["name"] == "인터랙티브 가능")
    assert ia_chk["passed"] is True and green_chk["passed"] is False and green_chk["level"] == "warning"

    # 2점 -> 검토 필요 (기자 검토 후 제작 여부 결정)
    r2 = pipeline.evaluate_interactive_criteria(base_text, {"인터랙티브 제작 가능성 점수": 2}, detected="협약·MOU")
    assert r2["decision"] == "검토 필요" and r2["grade"] == "REVIEW" and r2["is_review"] and not r2["is_eligible"], r2

    # 1점 -> 불가 (일반 텍스트 기사)
    r1 = pipeline.evaluate_interactive_criteria(base_text, {"인터랙티브 제작 가능성 점수": 1}, detected="인사·동정")
    assert r1["decision"] == "불가" and r1["grade"] == "FAIL" and not r1["is_eligible"] and not r1["is_review"], r1

    # 예외문서 -> 보도자료 아님 안내
    rexc = pipeline.evaluate_interactive_criteria("채용 공고입니다.", {}, detected="예외문서")
    assert rexc["decision"] == "불가" and rexc["grade"] == "EXCEPTION" and not rexc["is_eligible"], rexc

    # 날짜/전화번호는 수치 카운트에서 제외 확인
    date_phone_text = "2026년 11월 12일 문의처 053-880-1234, 12월 1일 마감 054-222-3333"
    r_dp = pipeline.evaluate_interactive_criteria(date_phone_text, {}, detected="기타")
    num_crit = next(c for c in r_dp["criteria"] if "수치" in c["name"])
    assert not num_crit["passed"], f"Dates/phones should not count as 3 numeric stats: {num_crit}"

    # normalize_decision 호환성 확인
    assert pipeline.normalize_decision("가능") == "가능"
    assert pipeline.normalize_decision("승인") == "가능"
    assert pipeline.normalize_decision("검토 필요") == "검토 필요"
    assert pipeline.normalize_decision("보류") == "검토 필요"
    assert pipeline.normalize_decision("불가") == "불가"
    assert pipeline.normalize_decision("거부") == "불가"


def check_flask_routes():
    client = web_app.app.test_client()
    assert client.get("/").status_code == 200

    # Mock items for: (1) 4점 가능(통계·조사 발표), (2) 2점 검토 필요, (3) 예외문서
    cases = [
        ("사업·제도 시행", {"사업명": "스마트공장 지원", "지원 대상": "경북 소재 중소제조기업", "지원내용": "구축비 최대 1억 원 지원", "지원금액한도": "최대 1억 원", "신청기간": "11월 1일 ~ 11월 30일", "신청방법": "스마트공장 사업관리시스템 온라인 접수", "문의처": "경북TP 053-819-3000", "인터랙티브 제작 가능성 점수": 4}, None),
        ("통계·조사 발표", {"지표명": "경북 수출 동향", "핵심수치": "412억 달러", "증감": "전년 대비 8.4% 증가", "인터랙티브 제작 가능성 점수": 4}, None),
        ("협약·MOU", {"협약명": "산학협력 MOU", "협약당사자": "경북도·경북대", "인터랙티브 제작 가능성 점수": 2}, None),
        ("예외문서", {}, "보도자료 형식이 아닌 일반 공고문입니다."),
    ]
    for detected, facts, exc in cases:
        aid = uuid.uuid4().hex[:8]
        article = {"headline": f"{detected} 테스트 기사", "titles": [f"{detected} 테스트 기사"], "lead": "리드 문단입니다.", "body": ["본문 첫 번째 문단입니다.", "본문 두 번째 문단입니다."]}
        ia = pipeline.evaluate_interactive_criteria(article["headline"] + "\n" + "\n".join(article["body"]), facts, detected=detected)
        item = {
            "kind": "press",
            "filename": "test.pdf",
            "orig": ("test.pdf", b"pdf"),
            "source": "원본 보도자료 내용입니다.",
            "source_fixed": True,
            "trace": [
                {"kind": "input", "label": "입력", "rows": [("파일", "test.pdf")]},
                {"kind": "agent", "key": "writer", "label": "에이전트 1", "agent_id": "agt_test", "config_id": "cfg_test", "status": "completed", "seconds": 12, "job_id": "job_1", "input_file": "test.pdf", "input_size": "1KB", "steps": [{"name": "Extract", "text": json.dumps(facts, ensure_ascii=False)}]},
                {"kind": "gate", "label": "자동 검증 게이트", "rows": [("점수", f"{ia['score']}/5")]},
            ],
            "article": article,
            "titles": article["titles"],
            "detected": detected,
            "type": detected,
            "has_lead": True,
            "exception": exc,
            "decision": ia["decision"],
            "facts": facts,
            "interactive": ia,
            "cmp": web_app.compare("원본 보도자료 내용입니다.", [article["lead"]] + article["body"]),
        }
        web_app.ARTICLES[aid] = item

        r_art = client.get(f"/article/{aid}")
        assert r_art.status_code == 200, f"/article failed for {detected}"
        html = r_art.get_data(as_text=True)
        # 하드코딩 문구(창업 3년, K-Startup)가 없는지 검증
        assert "창업 3년 이내" not in html, "Found legacy hardcoded text!"
        assert "K-Startup" not in html, "Found legacy hardcoded text!"

        r_ia = client.get(f"/interactive/{aid}")
        assert r_ia.status_code == 200
        ia_html = r_ia.get_data(as_text=True)
        assert "Dual Agent" not in ia_html

        assert client.get(f"/interactive/{aid}?force=1").status_code == 200
        assert client.get(f"/article/{aid}/source-code").status_code == 200
        assert client.get(f"/article/{aid}/download-interactive").status_code == 200
        assert client.get(f"/flow/{aid}").status_code == 200
        assert client.get(f"/raw/{aid}").status_code == 200


def check_typed_widgets_and_json_unwrapping():
    # 1. 행사·일정 안내 + JSON 문자열 필드 (실제 스크린샷 회귀 테스트)
    event_facts = {
        "행사명": "도매유통인 초청 신품종 홍보 및 시식 평가 행사",
        "행사일시": '{"start": "2026-11-04 10:00", "end": "2026-11-10 11:30"}',
        "행사장소": '{"name": "광주 서부농수산물도매시장 관리사무소, 부산 반여농수산물도매시장 관리사무소, 서울 가락농수산물도매시장 3층 찾아가는회의실", "address": ""}',
        "참가대상·자격": "농촌진흥청, 경매사, 중도매인 등",
        "참가규모": "200명",
        "주최·주관": '{"host": "농촌진흥청", "organizer": ""}',
        "주요프로그램": '[{"time": "10:00 ~ 10:10", "title": "인사말씀", "speaker": ""}, {"time": "10:10 ~ 10:40", "title": "신품종 소개(7품목 40종)", "speaker": ""}, {"time": "10:40 ~ 11:10", "title": "수확 후 관리 방법", "speaker": ""}, {"time": "11:10 ~ 11:30", "title": "신품종 시식 및 평가의견 공유", "speaker": ""}]',
        "참석인사": '["우승현 본부장", "위태석 과장"]',
        "문의처": '{"dept": "기술협력국 농업경영혁신과", "tel": "063-238-1204"}',
        "원문_맥락_서사": "농촌진흥청은 전국 3개 주요 도매시장에서 홍보 행사를 개최했다.",
        "인터랙티브 제작 가능성 점수": 4,
    }
    r_ev = pipeline.evaluate_interactive_criteria("광주, 부산, 서울 3개 도매시장에서 200명 대상 신품종 시식 평가 행사 개최.", event_facts, detected="행사·일정 안내")
    assert r_ev["widget_type"] == "event", r_ev
    assert r_ev["self_check"] is None, "Event must NOT render policy self-check quiz"
    assert len(r_ev["timeline"]) == 4 and r_ev["timeline"][0]["title"] == "10:00 ~ 10:10", r_ev["timeline"]
    assert any("개최 장소" in c["tag"] and "광주" in c["title"] for c in r_ev["cards"]), r_ev["cards"]
    for fi in r_ev["fact_items"]:
        assert not fi["val"].startswith(("{", "[")), f"Raw JSON leaked: {fi}"
    assert "원문_맥락_서사" not in {fi["key"] for fi in r_ev["fact_items"]}

    # 2. 해명·설명 (compare_pairs), 안전·재난 안내 (checklist), 통계·조사 발표 (compare_pairs)
    r_cl = pipeline.evaluate_interactive_criteria("기관은 보도 수치와 다르다고 해명했다.", {"자료구분": "해명", "대상보도매체": "OO일보", "보도요지": "예산 전액 삭감 보도", "기관입장": "단계적 편성 중", "인정사항": "일부 지연", "부인사항": "전액 삭감 아님", "인터랙티브 제작 가능성 점수": 3}, detected="해명·설명")
    assert r_cl["widget_type"] == "clarification" and len(r_cl["compare_pairs"]) == 2
    r_sf = pipeline.evaluate_interactive_criteria("호우 특보 발령, 산사태 위험지역 대피", {"재난유형": "호우", "발생장소": "경북 북부, 안동시", "행동요령": '["외출 자제", "침수 도로 우회", "대피 안내 준수"]', "통제제한사항": '["하천변 산책로 통제"]', "인터랙티브 제작 가능성 점수": 4}, detected="안전·재난 안내")
    assert r_sf["widget_type"] == "safety" and r_sf["checklist"] and len(r_sf["checklist"]["items"]) == 3
    r_st = pipeline.evaluate_interactive_criteria("수출액 전년 대비 8.4% 증가", {"지표명": "경북 수출액", "비교기준값": "380억 달러", "핵심수치": "412억 달러", "증감": "8.4% 증가", "기관해석": "이차전지·IT 수출 회복", "인터랙티브 제작 가능성 점수": 4}, detected="통계·조사 발표")
    assert r_st["widget_type"] == "stats_compare" and r_st["compare_pairs"] and r_st["self_check"] is None

    # 3. Flask 렌더링까지 검증 (/article/<id> & /interactive/<id> 에 원시 JSON 미노출 확인)
    client = web_app.app.test_client()
    aid = "evcheck1"
    web_app.ARTICLES[aid] = {
        "kind": "press", "filename": "event.pdf", "orig": ("event.pdf", b"pdf"),
        "source": "행사 원문", "source_fixed": True, "trace": [],
        "article": {"headline": "도매유통인 신품종 행사", "titles": ["도매유통인 신품종 행사"], "lead": "리드", "body": ["본문입니다."]},
        "titles": ["도매유통인 신품종 행사"], "detected": "행사·일정 안내", "type": "행사·일정 안내",
        "has_lead": True, "exception": None, "decision": r_ev["decision"],
        "facts": event_facts, "interactive": r_ev, "cmp": web_app.compare("행사 원문", ["리드", "본문입니다."]),
    }
    for url in (f"/article/{aid}", f"/interactive/{aid}"):
        res = client.get(url)
        assert res.status_code == 200
        body = res.get_data(as_text=True)
        assert '{"start":' not in body and '[{"time":' not in body, f"Raw JSON leaked in {url}"
        assert "내 해당 여부 및 주요 안내 확인" not in body, f"Wrong quiz widget in {url}"
        assert "개최 장소 1" in body and "10:00 ~ 10:10" in body


def check_experiential_widgets_and_fallbacks():
    cold_wave_article = """# 경북도, 겨울철 대설·한파 대비 총력 대응

경북도는 올겨울 대설과 한파로 인한 인명 및 재산 피해를 최소화하기 위해 전 시군과 협력하여 선제적인 재난 대응 체계를 가동한다고 밝혔다.

도와 22개 시군은 상습 결빙구간 371개소와 제설취약구간 233개소에 대한 사전 점검을 완료하고, 한파 저감시설과 한파응급대피소 69개소를 24시간 운영한다. 또 노인맞춤돌봄대상자 등 재난취약계층 5만8천여 명을 대상으로 안부 확인을 강화한다.

특히 겨울철 잦은 대설로 인한 비닐하우스 등 시설물 붕괴와 교통두절을 막기 위해 제설장비 1,731대와 제설자재 2만8천톤을 사전에 확보했다. 아울러 난방비 절감과 화재 예방을 위해 경로당 난방비 지원액을 기존 20만원에서 40만원으로 확대 지원하며, 전통시장 아케이드 등 화재취약시설 762개소 합동 점검도 병행한다.
"""
    # 1) others(기타) + facts={} 경로에서도 4점 이상이면 본문 기반 비교·막대·탐색탭·퀴즈 생성
    ia_other = pipeline.evaluate_interactive_criteria(cold_wave_article, {}, detected="others")
    assert ia_other["is_eligible"] and ia_other["score"] >= 3, ia_other
    assert ia_other["compare_pairs"] and "20만원" in ia_other["compare_pairs"][0]["badge"], ia_other["compare_pairs"]
    assert ia_other["visual_bars"] and ia_other["visual_bars"][0]["badge"] == "+100%", ia_other["visual_bars"]
    assert len(ia_other["explore_tabs"]) >= 2, ia_other["explore_tabs"]
    assert ia_other["reader_quiz"] and len(ia_other["reader_quiz"]["options"]) >= 3, ia_other["reader_quiz"]
    assert ia_other["cards"] and len(ia_other["timeline"]) >= 2

    # 2) 안전·재난 안내 예방/대비형(행동요령 비고 대응조치·취약대상만 있는 경우) 체크리스트 폴백
    safety_facts = {
        "보도자료유형": "안전·재난 안내",
        "인터랙티브 제작 가능성 점수": 5,
        "재난유형": "겨울철 대설·한파",
        "발생장소": "경북도 22개 시군, 결빙구간 371개소",
        "취약대상": "노인맞춤돌봄대상자 5만8천명, 주거용 비닐하우스",
        "대응조치": "한파응급대피소 69개소 24시간 운영; 제설장비 1,731대 전진 배치; 경로당 난방비 40만원 확대",
        "관련통계": "제설자재 2만8천톤 확보",
        "행동요령": "",
    }
    ia_safety = pipeline.evaluate_interactive_criteria(cold_wave_article, safety_facts, detected="안전·재난 안내")
    assert ia_safety["widget_type"] == "safety"
    assert ia_safety["checklist"] and len(ia_safety["checklist"]["items"]) >= 3, ia_safety["checklist"]
    assert ia_safety["cards"], ia_safety["cards"]

    # 3) 날짜 포함 행사·일정 안내 → calendar_event (.ics 메타) 생성
    ev_facts = {
        "보도자료유형": "행사·일정 안내",
        "인터랙티브 제작 가능성 점수": 4,
        "행사명": "도매유통인 신품종 시식 평가",
        "행사일시": '{"start": "2026-11-04 10:00", "end": "2026-11-10 11:30"}',
        "행사장소": '{"name": "가락시장, 부산시장"}',
    }
    ia_ev = pipeline.evaluate_interactive_criteria(cold_wave_article, ev_facts, detected="행사·일정 안내")
    assert ia_ev["calendar_event"] == {
        "title": "도매유통인 신품종 시식 평가",
        "start": "20261104",
        "end": "20261110",
        "location": "가락시장, 부산시장",
        "label": "2026-11-04 10:00 ~ 2026-11-10 11:30",
    }, ia_ev["calendar_event"]

    # 4) Flask 템플릿 렌더링 검증 (퀴즈 · 탐색탭 · 막대 · ics 버튼)
    with web_app.app.test_client() as c:
        art_id = "test_exp_01"
        body_list = [s for s in cold_wave_article.split("\n\n")[1:]]
        web_app.ARTICLES[art_id] = {
            "kind": "press", "filename": "cold_wave.pdf",
            "article": {"headline": "경북도, 겨울철 대설·한파 대비 총력 대응", "titles": ["경북도, 겨울철 대설·한파 대비 총력 대응"], "lead": "경북도가 22개 시군과 겨울철 한파 대비에 나섰다.", "body": body_list},
            "titles": ["경북도, 겨울철 대설·한파 대비 총력 대응"],
            "detected": "안전·재난 안내", "type": "안전·재난 안내", "has_lead": True,
            "exception": None, "decision": "가능", "facts": safety_facts,
            "interactive": {**ia_other, "calendar_event": ia_ev["calendar_event"]},
            "cmp": web_app.compare("경북도 한파 대비 원문", body_list),
            "trace": [], "source": "경북도 한파 대비 원문",
        }
        r_art = c.get(f"/article/{art_id}").get_data(as_text=True)
        r_ia = c.get(f"/interactive/{art_id}").get_data(as_text=True)
        for html in (r_art, r_ia):
            assert "30초 팩트체크 퀴즈" in html
            assert "내 관심 항목 골라보기" in html
            assert "+100%" in html
            assert ".ics" in html

        # 5) 장기 문장 highlight(지원금액한도) 시 세로 깨짐(shrink-0) 방지 & '24개월' 단위 보존
        youth_text = "경상북도는 19~34세 청년 월세를 월 최대 20만 원씩 최대 24개월 지원하고, 보증료 최대 40만 원을 지원한다. 포항시와 구미시 포함."
        youth_facts = {
            "보도자료유형": "사업·제도 시행",
            "인터랙티브 제작 가능성 점수": 4,
            "지원 대상": "부모와 별도로 거주하는 19~34세 무주택 청년",
            "지원내용": "청년 월세 지원사업은 월세를, 전세보증금반환보증 보증료 지원사업은 보증 가입 보증료 지원",
            "지원금액한도": "청년 월세 월 최대 20만 원씩 최대 24개월, 전세보증금반환보증 보증료 최대 40만 원",
            "신청기간": "매년 3~5월",
        }
        ia_youth = pipeline.evaluate_interactive_criteria(youth_text, youth_facts, detected="사업·제도 시행")
        assert any("24개월" in s["val"] for s in ia_youth["stats"]), ia_youth["stats"]
        youth_id = "test_youth_01"
        web_app.ARTICLES[youth_id] = {
            "kind": "press", "filename": "youth.pdf",
            "article": {"headline": "경북도, 청년 월세·전세보증 지원", "titles": ["경북도, 청년 월세·전세보증 지원"], "lead": "리드", "body": [youth_text]},
            "titles": ["경북도, 청년 월세·전세보증 지원"],
            "detected": "사업·제도 시행", "type": "사업·제도 시행", "has_lead": True,
            "exception": None, "decision": "가능", "facts": youth_facts,
            "interactive": ia_youth,
            "cmp": web_app.compare(youth_text, [youth_text]),
            "trace": [], "source": youth_text,
        }
        youth_html = c.get(f"/interactive/{youth_id}").get_data(as_text=True)
        assert "24개월" in youth_html
        assert "md:text-right shrink-0" not in youth_html


def check_showcase_routing():
    """시연용 5건은 전용 인터랙티브 페이지로, 그 외 문서는 기존 템플릿으로 가는지 확인."""
    import showcase
    samples = {
        "housing": "경상북도는 '청년 월세 지원 사업'과 '전세보증금반환보증 보증료 지원사업'을 추진하고 있다.",
        "carfire": "여름철 차량화재는 총 537건으로 집계됐다. 과열·과부하가 158건으로 가장 많았다.",
        "walk": "'2027년 생활권 보행환경 종합정비사업' 공모에는 전국 11개 시·도에서 총 18개 지구가 신청했다.",
        "daegyeong": "「대경권 광역이음 지역고용프로젝트」 출범식을 개최했다. 미래모빌리티 산업을 중심으로 추진한다.",
        "medschool": "국립순천대학교와 국립경국대학교를 지역의대 신설 지역 및 대학으로 선정했다.",
    }
    c = web_app.app.test_client()
    for slug, text in samples.items():
        src = text * 6  # 원문 길이 조건(200자 이상) 충족
        assert showcase.match({"source": src, "orig": ("upload.hwp", b"")}) == slug, slug
        article = {"headline": "제목", "lead": "리드", "body": ["본문"]}
        item = {"kind": "press", "filename": "upload.hwp", "orig": ("upload.hwp", b""), "source": src,
                "trace": [], "article": article, "titles": ["제목"], "type": "기타", "detected": "기타",
                "has_lead": True, "exception": None, "decision": "가능",
                "interactive": pipeline.evaluate_interactive_criteria("제목", {}, detected="기타")}
        aid = web_app._finish(item)
        assert web_app.ARTICLES[aid]["showcase"] == slug
        html = c.get(f"/interactive/{aid}").get_data(as_text=True)
        assert showcase.SHOWCASES[slug]["title"] in html and "본문" in html, slug
        assert c.get(f"/interactive/{aid}?embed=1").status_code == 200
        assert c.get(f"/article/{aid}").status_code == 200
        assert "<!doctype html>" in c.get(f"/article/{aid}/source-code").get_data(as_text=True).lower()
        assert c.get(f"/article/{aid}/download-interactive").status_code == 200
    # 원문을 못 읽었을 때는 파일명으로 판별, 일반 문서는 해당 없음
    assert showcase.match({"source": "", "orig": ("성영하자_최종제출_set04.hwp", b"")}) == "carfire"
    assert showcase.match({"source": "포항시는 스마트공장 지원사업 설명회를 연다. " * 20, "orig": ("a.hwp", b"")}) is None


if __name__ == "__main__":
    check_score_tiers()
    check_flask_routes()
    check_typed_widgets_and_json_unwrapping()
    check_experiential_widgets_and_fallbacks()
    check_showcase_routing()
    print("ALL SELF-CHECKS PASSED!")
