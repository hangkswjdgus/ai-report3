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
        tone = data.get("tone") or "자연스럽게"

        purpose = data.get("purpose") or "일반적인 설명"
        details = data.get("details") or "기본적인 내용"
        extra = data.get("extra") or "없음"

        print("요청 들어옴:", topic)

        if template == "report":
            prompt = f"""
주제: {topic}
목적: {purpose}
내용: {details}
추가: {extra}

보고서를 작성해라.
길이:{length}
문체:{tone}
"""

        elif template == "paper":
            prompt = f"""
주제:{topic}
연구 목적:{purpose}
내용:{details}

논문 형식으로 작성
"""

        elif template == "resume":
            prompt = f"""
주제:{topic}
내용:{details}

이력서 작성
"""

        else:
            prompt = f"주제:{topic} 발표문 작성"

        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role":"user","content":prompt}],
            request_timeout=40
        )

        return jsonify({
            "result": response["choices"][0]["message"]["content"]
        })

    except Exception as e:
        print("에러:", e)
        return jsonify({
            "result": f"오류 발생: {str(e)}"
        })


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)