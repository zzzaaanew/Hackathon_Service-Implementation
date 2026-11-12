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


if __name__ == "__main__":
    check_score_tiers()
    check_flask_routes()
    print("ALL SELF-CHECKS PASSED!")
