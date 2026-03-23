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
        length_rule = "본문은 최소 800자 이상으로 작성할 것."
    elif length == "길게":
        length_rule = "본문은 최소 1800자 이상으로 작성할 것."
    else:
        length_rule = "본문은 최소 1200자 이상으로 작성할 것."

    ccommon_rules = f"""
너는 대학 과제, 보고서, 발표문, 논문 초안을 작성하는 문서 생성 보조 시스템이다.

절대 규칙:

1. 출처가 불분명한 통계, 수치, 연구명, 기관명, 논문명, 퍼센트, 연도는 절대 지어내지 말 것.
2. "연구에 따르면", "조사에 따르면" 같은 표현을 사용할 경우 반드시 실제 수치와 함께 써야 한다.
3. 수치(%, 증가율, 감소율, 인원, 규모 등)가 없는 주장은 금지한다.
4. 수치가 없는 경우에는 반드시 다음 형식으로 작성할 것:
   → "○○기관의 조사 자료에 따르면 ~"
   → "○○ 연구에서는 ~ 경향이 확인되었다"
5. 절대로 애매한 표현 금지:
   - "많은 연구에서"
   - "일부 연구에서"
   - "전반적으로"
   - "대체로 증가했다"
6. 가능한 경우 아래 형식으로 작성:
   → "2023년 교육부 조사에 따르면 AI 활용 수업 참여 학생의 68%가 학습 이해도가 향상되었다고 응답하였다."
7. 수치를 쓸 때는 반드시:
   - 연도
   - 기관 또는 연구 출처
   - 구체적인 결과
   를 포함해야 한다.
8. {length_rule}
9. 문체는 {tone}로 작성할 것.
10. 문서 구조를 제목과 소제목으로 명확히 나눌 것.
11. 마지막 참고자료는 실제 사용한 자료만 정리할 것.
12. 반드시 JSON 형식으로만 답할 것.
13. 설명 없이 JSON만 출력할 것.
14. 동일한 문장에서 "연구에 따르면" 표현을 반복하지 말 것.
15. 각 문단마다 최소 1개 이상의 구체적 근거(수치 또는 기관)를 포함할 것.

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

- 반드시 웹 검색을 통해 실제로 존재하는 통계 또는 연구 자료를 우선 사용해라.
- 수치가 포함된 자료를 우선적으로 선택해라.
- 가능하면 정부기관, 대학, 학술지, 공공기관 자료를 사용해라.
- 한국 자료가 있으면 우선 사용해라.
- references에는 실제 사용한 자료만 넣어라.
- 최소 1개 이상의 "숫자 포함 근거"를 반드시 포함할 것.
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