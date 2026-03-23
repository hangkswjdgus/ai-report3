from flask import Flask, render_template, request, jsonify
from openai import OpenAI
import os
import json
import re
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    raise ValueError("OPENAI_API_KEY 환경변수가 설정되지 않았습니다.")

client = OpenAI(api_key=api_key)


def build_prompt(template, topic, length, tone, purpose, details, extra):
    if length == "짧게":
        length_rule = "본문은 최소 1600자 이상으로 작성할 것."
    elif length == "길게":
        length_rule = "본문은 최소 3600자 이상으로 작성할 것."
    else:
        length_rule = "본문은 최소 2400자 이상으로 작성할 것."

    common_rules = f"""
너는 대학 과제, 보고서, 발표문, 논문 초안을 작성하는 문서 생성 보조 시스템이다.

절대 규칙:
1. 출처가 불분명한 통계, 수치, 연구명, 기관명, 논문명, 퍼센트, 연도는 절대 지어내지 말 것.
2. 핵심 주장에는 가능한 한 실제 자료, 통계, 연구 결과, 기관 보고서 등의 근거를 포함할 것.
3. 다만 보고서는 단순히 근거를 나열하는 방식이 아니라, 근거를 바탕으로 해석과 주장까지 충분히 전개할 것.
4. 숫자나 통계가 있는 경우에는 반드시 그 의미를 설명할 것.
   예: "이 수치는 AI 교육이 단순한 관심 수준을 넘어 실제 학습 효과로 이어졌음을 보여준다."
5. 숫자가 없는 경우에도 기관명, 보고서명, 연구명 등을 바탕으로 내용을 설명할 수 있다. 단, 출처 없이 지어내면 안 된다.
6. "연구에 따르면", "조사에 따르면" 같은 표현을 쓸 때는 가능하면 구체적인 기관명, 연도, 수치 또는 연구 결과를 함께 제시할 것.
7. 각 문단은 아래 순서를 따르도록 할 것:
   - 주장 또는 핵심 내용 제시
   - 관련 근거 또는 자료 제시
   - 그 근거가 의미하는 바를 해석
   - 주제 전체와 연결되는 설명 추가
8. 각 문단을 짧게 끝내지 말고, 설명을 충분히 붙여서 서술형으로 전개할 것.
9. 단순한 정보 나열형 문장이 아니라, 대학 보고서처럼 문단 간 연결이 자연스럽고 논리 흐름이 이어지도록 작성할 것.
10. 한 문단 안에서 근거를 제시한 뒤에는 반드시 그에 대한 분석, 영향, 시사점, 한계, 또는 의미를 덧붙일 것.
11. 지나치게 딱딱하게 쓰지 말고, 제출용 보고서처럼 자연스럽고 설득력 있게 작성할 것.
12. 각 주요 소제목마다 최소 2개 이상의 충분한 설명 문단을 작성하려고 할 것.
13. {length_rule}
14. 문체는 {tone}로 작성할 것.
15. 문서 구조를 제목과 소제목으로 명확히 나눌 것.
16. 마지막 참고자료는 실제 사용한 자료만 정리할 것.
17. 반드시 JSON 형식으로만 답할 것.
18. 설명문, 코드블록, 안내문 없이 JSON 객체만 출력할 것.
19. body에는 문서 본문만 넣고, references에는 실제 사용한 참고자료만 넣을 것.
20. 너무 짧고 압축적인 요약문 대신, 근거와 해석이 함께 있는 풍부한 본문을 작성할 것.
21. 본문은 "근거 제시 → 의미 해석 → 주장 전개"의 흐름이 반복되도록 작성할 것.
22. 단락 사이 연결 문장을 넣어 글 전체가 끊기지 않게 할 것.

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

추가 작성 원칙:
- 서론에서는 주제의 중요성과 문제의식을 충분히 설명할 것.
- 본론에서는 단순 사례 나열이 아니라, 근거 제시 후 분석과 주장까지 이어갈 것.
- 결론에서는 전체 내용을 요약하는 데 그치지 말고, 최종적인 시사점이나 방향성까지 제시할 것.

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

추가 작성 원칙:
- 관련 연구에서는 기존 연구 내용을 단순 요약하지 말고, 어떤 공통점과 차이가 있는지도 설명할 것.
- 본문 분석에서는 자료를 해석하고, 연구적 관점의 주장도 분명히 제시할 것.

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

추가 작성 원칙:
- 단순 경력 나열이 아니라, 경험이 어떤 역량으로 이어졌는지 설명할 것.
- 주장 뒤에는 가능한 한 구체적인 사례를 붙이고, 사례 뒤에는 그 의미를 해석할 것.

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

추가 작성 원칙:
- 발표용 문서라도 메모 수준으로 짧게 쓰지 말고, 설명이 충분한 원고형 문장으로 작성할 것.
- 근거를 제시한 뒤에는 왜 중요한지와 어떤 메시지를 전달하는지까지 설명할 것.

사용자 입력:
주제: {topic}
발표 목적: {purpose}
핵심 내용: {details}
추가 요구사항: {extra}
"""

    web_rules = """
특히 중요:
- 반드시 웹 검색을 통해 실제로 확인 가능한 자료를 우선 사용해라.
- 한국 자료가 있으면 우선 사용하고, 부족하면 해외 공공기관, 대학, 학술 자료를 추가해라.
- 정부기관, 공공기관, 대학, 학술지, 국제기구 자료를 우선 사용해라.
- 수치가 포함된 자료가 있으면 우선 활용하되, 수치만 제시하고 끝내지 말고 그 의미를 설명해라.
- references에는 실제 사용한 자료만 넣어라.
- 최소 2개 이상의 신뢰할 수 있는 참고자료를 포함하려고 시도해라.
- 근거를 사용하되, 본문 전체가 근거 나열문처럼 보이지 않도록 설명과 해석을 충분히 덧붙여라.
"""

    return f"{common_rules}\n{structure}\n{web_rules}"


def extract_json_text(raw_text):
    """
    모델 응답에서 JSON 부분만 최대한 안정적으로 추출
    """
    text = raw_text.strip()

    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]

    if text.endswith("```"):
        text = text[:-3]

    text = text.strip()

    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        text = text[start:end + 1]

    return text.strip()


def safe_decode_json_string(value):
    """
    이미 정상 문자열이면 그대로 두고,
    JSON 이스케이프가 남아 있을 때만 최소한으로 정리
    """
    if not isinstance(value, str):
        return str(value)

    # 이미 정상 한글이면 절대 건드리지 않음
    if any('\uac00' <= ch <= '\ud7a3' for ch in value):
        return value

    # 줄바꿈 이스케이프만 최소한으로 정리
    value = value.replace("\\n", "\n")
    value = value.replace("\\t", "\t")
    value = value.replace('\\"', '"')

    return value


def parse_model_response(raw_text):
    cleaned_text = extract_json_text(raw_text)

    try:
        parsed = json.loads(cleaned_text)

        title = str(parsed.get("title", "생성된 문서")).strip()
        body = parsed.get("body", "")
        references = parsed.get("references", [])

        if not isinstance(body, str):
            body = str(body)

        if not isinstance(references, list):
            references = [str(references)]

        references = [str(ref).strip() for ref in references if str(ref).strip()]

        if not title:
            title = "생성된 문서"

        if not body:
            body = "본문을 받아오지 못했습니다."

        return title, body.strip(), references

    except Exception:
        title = "생성된 문서"
        body = raw_text.strip()
        references = ["출처 분리에 실패했습니다. 본문 내용을 직접 확인해 주세요."]

        try:
            title_match = re.search(r'"title"\s*:\s*"(.*?)"\s*,', raw_text, re.DOTALL)
            body_match = re.search(r'"body"\s*:\s*"(.*?)"\s*(,\s*"references"|})', raw_text, re.DOTALL)
            refs_match = re.search(r'"references"\s*:\s*(\[[\s\S]*?\])', raw_text, re.DOTALL)

            if title_match:
                title = title_match.group(1).replace("\\n", "\n").replace('\\"', '"').strip()

            if body_match:
                body = body_match.group(1).replace("\\n", "\n").replace('\\"', '"').strip()

            if refs_match:
                parsed_refs = json.loads(refs_match.group(1))
                if isinstance(parsed_refs, list):
                    references = [str(ref).strip() for ref in parsed_refs if str(ref).strip()]

        except Exception:
            pass

        return title, body, references


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json(silent=True) or {}

        template = (data.get("template") or "report").strip()
        topic = (data.get("topic") or "일반적인 주제").strip()
        length = (data.get("length") or "보통").strip()
        tone = (data.get("tone") or "자연스럽고 이해하기 쉽게").strip()
        purpose = (data.get("purpose") or "일반적인 설명").strip()
        details = (data.get("details") or "기본적인 내용").strip()
        extra = (data.get("extra") or "없음").strip()

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
        title, body, references = parse_model_response(raw_text)

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