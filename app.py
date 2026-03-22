from flask import Flask, render_template, request, jsonify
import openai
import os
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
openai.api_key = os.getenv("OPENAI_API_KEY")


@app.route('/')
def home():
    return render_template("index.html")


@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.json or {}

        template = data.get("template", "report")
        topic = data.get("topic") or "일반적인 주제"

        length = data.get("length") or "보통"
        tone = data.get("tone") or "자연스럽고 이해하기 쉽게"

        purpose = data.get("purpose") or "일반적인 설명"
        details = data.get("details") or "기본적인 내용"
        extra = data.get("extra") or "없음"

        # 🔥 길이 강제 설정
        if length == "짧게":
            length_instruction = "최소 600자 이상"
        elif length == "보통":
            length_instruction = "최소 1000자 이상"
        else:
            length_instruction = "최소 1500자 이상"

        # 🔥 공통 요구사항 (핵심🔥)
        common_rules = f"""
반드시 지켜야 할 조건:
- {length_instruction}로 작성할 것
- 내용은 구체적으로 작성할 것
- 주장에는 반드시 근거를 포함할 것
- 통계 자료 또는 연구 사례를 포함할 것
- 단순 설명이 아닌 분석 중심으로 작성할 것
"""

        # 🔥 템플릿별 프롬프트
        if template == "report":
            prompt = f"""
다음 정보를 바탕으로 대학 과제용 보고서를 작성하라.

주제: {topic}
작성 목적: {purpose}
포함 내용: {details}
추가 요구사항: {extra}

구조:
1. 서론 - 주제 소개 및 중요성 설명
2. 본론 - 핵심 내용, 근거, 사례, 통계 포함
3. 결론 - 요약 및 개인 의견

{common_rules}

문체: {tone}
"""

        elif template == "paper":
            prompt = f"""
다음 정보를 바탕으로 학술 논문 형식으로 작성하라.

주제: {topic}
연구 목적: {purpose}
연구 내용: {details}
추가 요구사항: {extra}

구조:
1. 서론
2. 관련 연구
3. 방법론
4. 결과 및 분석
5. 결론

{common_rules}

문체: 매우 공식적이고 전문적으로 작성할 것
"""

        elif template == "resume":
            prompt = f"""
다음 정보를 바탕으로 자기소개서를 작성하라.

주제: {topic}
지원 목적: {purpose}
경험: {details}
추가 요구사항: {extra}

구조:
- 지원 동기
- 경험 및 역량
- 강점
- 마무리

{common_rules}

문체: {tone}
"""

        elif template == "presentation":
            prompt = f"""
다음 정보를 바탕으로 발표용 원고를 작성하라.

주제: {topic}
발표 목적: {purpose}
핵심 내용: {details}
추가 요구사항: {extra}

구조:
- 도입
- 핵심 내용 설명
- 사례 및 근거
- 결론

{common_rules}

문체: 발표용으로 명확하고 이해하기 쉽게 작성
"""

        else:
            prompt = f"""
주제: {topic}
내용을 작성하라.
{common_rules}
"""

        # 🔥 OpenAI 호출
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            request_timeout=60
        )

        result_text = response["choices"][0]["message"]["content"]

        return jsonify({
            "result": result_text
        })

    except Exception as e:
        return jsonify({
            "result": f"오류 발생: {str(e)}"
        })


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)