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
        data = request.json

        template = data.get("template")
        topic = data.get("topic")
        length = data.get("length")
        tone = data.get("tone")
        purpose = data.get("purpose")
        details = data.get("details")
        extra = data.get("extra")

        # 🔥 템플릿별 프롬프트 생성
        if template == "report":
            prompt = f"""
            다음 정보를 바탕으로 대학 보고서를 작성해라.

            주제: {topic}
            작성 목적: {purpose}
            포함 내용: {details}
            추가 요구사항: {extra}

            조건:
            - 길이: {length}
            - 문체: {tone}
            - 서론, 본론, 결론 구조
            - 논리적이고 자연스럽게 작성
            """

        elif template == "paper":
            prompt = f"""
            다음 정보를 바탕으로 학술 논문 형식으로 작성해라.

            주제: {topic}
            연구 목적: {purpose}
            연구 내용: {details}
            추가 요구사항: {extra}

            조건:
            - 길이: {length}
            - 문체: 공식적이고 전문적으로
            - 서론, 관련 연구, 방법론, 결과, 결론 포함
            """

        elif template == "resume":
            prompt = f"""
            다음 정보를 바탕으로 자기소개서 또는 이력서를 작성해라.

            주제: {topic}
            지원 목적: {purpose}
            경험 및 내용: {details}
            추가 요구사항: {extra}

            조건:
            - 길이: {length}
            - 문체: {tone}
            - 강점이 잘 드러나도록 작성
            """

        elif template == "presentation":
            prompt = f"""
            다음 정보를 바탕으로 발표용 원고를 작성해라.

            주제: {topic}
            발표 목적: {purpose}
            핵심 내용: {details}
            추가 요구사항: {extra}

            조건:
            - 길이: {length}
            - 문체: 발표용으로 간결하고 명확하게
            - 듣는 사람이 이해하기 쉽게 구성
            """

        else:
            prompt = f"""
            주제: {topic}
            내용을 작성해라.
            """

        # 🔥 OpenAI 호출
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}]
        )

        return jsonify({
            "result": response['choices'][0]['message']['content']
        })

    except Exception as e:
        return jsonify({
            "result": f"오류 발생: {str(e)}"
        }), 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)