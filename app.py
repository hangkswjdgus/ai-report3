from flask import Flask, render_template, request, jsonify
from openai import OpenAI
import os
import json
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))


def build_prompt(template, topic, length, tone, purpose, details, extra):
    if length == "짧게":
        length_rule = "본문은 최소 800자 이상으로 작성할 것."
    elif length == "길게":
        length_rule = "본문은 최소 1800자 이상으로 작성할 것."
    else:
        length_rule = "본문은 최소 1200자 이상으로 작성할 것."

    common_rules = f"""
너는 대학 과제, 보고서, 발표문, 논문 초안을 작성하는 문서 생성 보조 시스템이다.

절대 규칙:
1. 출처가 불분명한 통계, 수치, 연구명, 기관명, 논문명, 퍼센트, 연도는 절대 지어내지 말 것.
2. 핵심 주장에는 가능한 한 구체적인 출처를 포함할 것.
3. 수치가 있으면 반드시 출처와 함께 쓸 것.
4. 출처를 찾지 못한 내용은 "출처 확인 필요" 또는 "일반적 설명" 수준으로만 작성할 것.
5. 막연한 "한 연구에 따르면", "어느 나라에서는" 같은 표현은 금지한다.
6. {length_rule}
7. 문체는 {tone}로 작성할 것.
8. 문서 구조를 제목과 소제목으로 명확히 나눌 것.
9. 마지막 참고자료는 실제 사용한 자료만 정리할 것.
10. 반드시 JSON 형식으로만 답할 것.

반드시 아래 JSON 형식으로만 응답해라:
{{
  "title": "문서 제목",
  "body": "문서 본문 전체",
  "references": [
    "참고자료 1",
    "참고자료 2"
  ]
}}
"""

    if template == "report":
        structure = f"""
문서 유형: 대학 보고서

반드시 아래 구조를 본문(body)에 포함할 것:
# 서론
# 본론
# 결론

사용자 입력:
주제: {topic}
작성 목적: {purpose}
포함하고 싶은 내용: {details}
추가 요구사항: {extra}
"""
    elif template == "paper":
        structure = f"""
문서 유형: 논문 형식 초안

반드시 아래 구조를 본문(body)에 포함할 것:
# 초록
# 서론
# 관련 연구 또는 배경
# 본문 분석
# 결론

사용자 입력:
주제: {topic}
연구 목적: {purpose}
연구 내용: {details}
추가 요구사항: {extra}
"""
    elif template == "resume":
        structure = f"""
문서 유형: 자기소개서 / 이력서 보조문

반드시 아래 구조를 본문(body)에 포함할 것:
# 지원 동기
# 경험 및 역량
# 강점
# 마무리

사용자 입력:
주제: {topic}
지원 목적: {purpose}
경험 및 내용: {details}
추가 요구사항: {extra}
"""
    else:
        structure = f"""
문서 유형: 발표문

반드시 아래 구조를 본문(body)에 포함할 것:
# 도입
# 핵심 내용
# 근거 및 사례
# 결론

사용자 입력:
주제: {topic}
발표 목적: {purpose}
핵심 내용: {details}
추가 요구사항: {extra}
"""

    web_rules = """
특히 중요:
- 웹 검색을 통해 실제로 확인 가능한 근거만 사용해라.
- 한국 자료가 있으면 우선 포함하고, 부족하면 해외 공공기관/대학/학술 자료를 추가해라.
- 교육, 기술, 경제, 사회 주제는 정부기관, 공공기관, 대학, 학술지, 국제기구 자료를 우선 사용해라.
- references 배열에는 실제 사용한 출처만 넣어라.
- references는 최소 2개 이상 넣으려고 시도하되, 찾지 못하면 무리해서 지어내지 말아라.
"""

    return f"{common_rules}\n{structure}\n{web_rules}"


@app.route('/')
def home():
    return render_template("index.html")


@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.json or {}

        template = data.get("template") or "report"
        topic = data.get("topic") or "일반적인 주제"
        length = data.get("length") or "보통"
        tone = data.get("tone") or "자연스럽고 이해하기 쉽게"
        purpose = data.get("purpose") or "일반적인 설명"
        details = data.get("details") or "기본적인 내용"
        extra = data.get("extra") or "없음"

        prompt = build_prompt(
            template=template,
            topic=topic,
            length=length,
            tone=tone,
            purpose=purpose,
            details=details,
            extra=extra
        )

        response = client.responses.create(
            model="gpt-4.1-mini",
            tools=[{"type": "web_search_preview"}],
            input=prompt
        )

        raw_text = response.output_text.strip()

        try:
            parsed = json.loads(raw_text)
            title = parsed.get("title", "생성된 문서")
            body = parsed.get("body", "")
            references = parsed.get("references", [])
        except Exception:
            # JSON 파싱 실패 시 안전한 fallback
            title = "생성된 문서"
            body = raw_text
            references = ["출처 분리에 실패했습니다. 본문 내용을 직접 확인해 주세요."]

        return jsonify({
            "title": title,
            "body": body,
            "references": references
        })

    except Exception as e:
        return jsonify({
            "title": "오류",
            "body": f"오류 발생: {str(e)}",
            "references": []
        }), 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)