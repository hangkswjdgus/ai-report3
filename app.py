from flask import Flask, render_template, request, jsonify
import openai
import os
from dotenv import load_dotenv

# .env 파일 불러오기
load_dotenv()

app = Flask(__name__)

# 환경변수에서 API 키 가져오기
openai.api_key = os.getenv("OPENAI_API_KEY")

@app.route('/')
def home():
    return render_template("index.html")

@app.route('/generate', methods=['POST'])
def generate():
    try:
        topic = request.json['topic']

        prompt = f"""
        주제: {topic}
        대학 보고서 형식으로 작성해줘.
        서론, 본론, 결론 포함.
        """

        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}]
        )

        return jsonify({
            "result": response['choices'][0]['message']['content']
        })

    except Exception as e:
        return jsonify({
            "result": f"오류가 발생했습니다: {str(e)}"
        }), 500

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)