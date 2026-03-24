from flask import Flask, render_template, request, jsonify, send_file
from openai import OpenAI
from dotenv import load_dotenv
import os
import json
import re
from io import BytesIO
from typing import Any, Dict, List, Tuple
from datetime import datetime

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

load_dotenv()

app = Flask(__name__)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
MODEL_NAME = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")

if not OPENAI_API_KEY:
    raise ValueError("OPENAI_API_KEY 환경변수가 설정되지 않았습니다.")

client = OpenAI(api_key=OPENAI_API_KEY)

ALLOWED_TEMPLATES = {"report", "paper", "resume", "presentation"}
ALLOWED_LENGTHS = {"짧게", "보통", "길게"}
ALLOWED_MODES = {
    "natural",
    "draft_assist",
    "revision_assist",
    "evidence_boost",
    "personalized",
}
ALLOWED_REF_STYLES = {"default", "apa", "mla", "chicago"}


def normalize_value(value: Any, default: str = "") -> str:
    if value is None:
        return default
    return str(value).strip()


def sanitize_input(data: Dict[str, Any]) -> Dict[str, str]:
    template = normalize_value(data.get("template"), "report")
    if template not in ALLOWED_TEMPLATES:
        template = "report"

    length = normalize_value(data.get("length"), "보통")
    if length not in ALLOWED_LENGTHS:
        length = "보통"

    mode = normalize_value(data.get("mode"), "natural")
    if mode not in ALLOWED_MODES:
        mode = "natural"

    ref_style = normalize_value(data.get("reference_style"), "apa").lower()
    if ref_style not in ALLOWED_REF_STYLES:
        ref_style = "apa"

    topic = normalize_value(data.get("topic"), "일반적인 주제")
    tone = normalize_value(data.get("tone"), "자연스럽고 이해하기 쉽게")
    purpose = normalize_value(data.get("purpose"), "일반적인 설명")
    details = normalize_value(data.get("details"), "기본적인 내용")
    extra = normalize_value(data.get("extra"), "없음")
    user_perspective = normalize_value(data.get("user_perspective"), "")
    class_context = normalize_value(data.get("class_context"), "")
    personal_experience = normalize_value(data.get("personal_experience"), "")
    professor_style = normalize_value(data.get("professor_style"), "")
    major = normalize_value(data.get("major"), "")
    target_company = normalize_value(data.get("target_company"), "")

    intro_weight = normalize_value(data.get("intro_weight"), "보통")
    body_weight = normalize_value(data.get("body_weight"), "보통")
    conclusion_weight = normalize_value(data.get("conclusion_weight"), "보통")
    balance_view = normalize_value(data.get("balance_view"), "일반")
    include_my_opinion = normalize_value(data.get("include_my_opinion"), "아니오")
    critical_view = normalize_value(data.get("critical_view"), "아니오")

    return {
        "template": template,
        "topic": topic,
        "length": length,
        "tone": tone,
        "purpose": purpose,
        "details": details,
        "extra": extra,
        "mode": mode,
        "user_perspective": user_perspective,
        "class_context": class_context,
        "personal_experience": personal_experience,
        "professor_style": professor_style,
        "major": major,
        "target_company": target_company,
        "reference_style": ref_style,
        "intro_weight": intro_weight,
        "body_weight": body_weight,
        "conclusion_weight": conclusion_weight,
        "balance_view": balance_view,
        "include_my_opinion": include_my_opinion,
        "critical_view": critical_view,
    }


def get_length_rule(length: str) -> str:
    if length == "짧게":
        return "본문은 최소 1600자 이상으로 작성할 것."
    elif length == "길게":
        return "본문은 최소 3200자 이상으로 작성할 것."
    return "본문은 최소 2200자 이상으로 작성할 것."


def get_mode_rules(mode: str) -> str:
    if mode == "natural":
        return """
작성 모드: 자연스러운 문체 모드
- 문장을 부드럽고 자연스럽게 연결할 것.
- 과하게 기계적인 반복 표현은 줄일 것.
"""
    if mode == "draft_assist":
        return """
작성 모드: 초안 보조 모드
- 구조를 분명히 잡고 수정하기 쉬운 형태로 작성할 것.
"""
    if mode == "revision_assist":
        return """
작성 모드: 수정 보조 모드
- 문단 간 연결과 표현 정리를 중점적으로 개선할 것.
"""
    if mode == "evidence_boost":
        return """
작성 모드: 근거 강화 모드
- 핵심 주장마다 가능한 한 구체적인 통계, 보고서, 논문, 기관 자료를 반영할 것.
"""
    if mode == "personalized":
        return """
작성 모드: 개인화 모드
- 사용자가 제공한 관점, 맥락, 경험, 전공 정보를 적극 반영할 것.
- 제공되지 않은 개인 경험은 지어내지 말 것.
"""
    return ""


def get_reference_style_instruction(style: str) -> str:
    if style == "apa":
        return "참고문헌은 APA 스타일에 최대한 가깝게 정리할 것."
    if style == "mla":
        return "참고문헌은 MLA 스타일에 최대한 가깝게 정리할 것."
    if style == "chicago":
        return "참고문헌은 Chicago 스타일에 최대한 가깝게 정리할 것."
    return "참고문헌은 읽기 좋은 기본 학술 형식으로 정리할 것."


def get_assignment_rules(payload: Dict[str, str]) -> str:
    return f"""
과제 특화 옵션:
- 서론 비중: {payload["intro_weight"]}
- 본론 비중: {payload["body_weight"]}
- 결론 비중: {payload["conclusion_weight"]}
- 찬반 균형: {payload["balance_view"]}
- 내 의견 포함: {payload["include_my_opinion"]}
- 비판적 관점 포함: {payload["critical_view"]}
"""


def get_template_structure(payload: Dict[str, str]) -> str:
    template = payload["template"]
    topic = payload["topic"]
    purpose = payload["purpose"]
    details = payload["details"]
    extra = payload["extra"]
    major = payload["major"]
    target_company = payload["target_company"]
    class_context = payload["class_context"]
    user_perspective = payload["user_perspective"]
    personal_experience = payload["personal_experience"]
    professor_style = payload["professor_style"]

    if template == "report":
        return f"""
문서 유형: 대학 보고서

반드시 아래 구조를 본문에 포함할 것:
# 서론
# 본론
# 결론

사용자 입력:
주제: {topic}
작성 목적: {purpose}
포함할 내용: {details}
추가 요구사항: {extra}
수업 맥락: {class_context}
사용자 관점: {user_perspective}
개인 경험 또는 생각: {personal_experience}
교수 스타일: {professor_style}
전공: {major}
"""

    if template == "paper":
        return f"""
문서 유형: 논문 형식 초안

반드시 아래 구조를 본문에 포함할 것:
# 초록
# 서론
# 관련 연구 또는 이론적 배경
# 본문 분석
# 결론

사용자 입력:
주제: {topic}
연구 목적: {purpose}
연구 내용: {details}
추가 요구사항: {extra}
수업 맥락: {class_context}
사용자 관점: {user_perspective}
교수 스타일: {professor_style}
전공: {major}
"""

    if template == "resume":
        return f"""
문서 유형: 자기소개서 / 이력서 보조문

반드시 아래 구조를 본문에 포함할 것:
# 지원 동기
# 경험 및 역량
# 강점
# 마무리

사용자 입력:
주제: {topic}
지원 목적: {purpose}
경험 및 내용: {details}
추가 요구사항: {extra}
개인 경험 또는 생각: {personal_experience}
사용자 관점: {user_perspective}
지원 대상 기업/기관: {target_company}
전공: {major}
"""

    return f"""
문서 유형: 발표문

반드시 아래 구조를 본문에 포함할 것:
# 도입
# 핵심 내용
# 근거 및 사례
# 결론

사용자 입력:
주제: {topic}
발표 목적: {purpose}
핵심 내용: {details}
추가 요구사항: {extra}
수업 맥락: {class_context}
사용자 관점: {user_perspective}
개인 경험 또는 생각: {personal_experience}
전공: {major}
"""


def build_prompt(payload: Dict[str, str]) -> str:
    reference_style = payload["reference_style"]

    return f"""
너는 상업용 문서 생성 서비스의 문서 작성 엔진이다.
사용자 입력을 바탕으로 완성도 높은 한국어 문서를 작성하라.

핵심 규칙:
1. 본문에는 인용을 자연스럽게 포함할 수 있다.
2. 참고문헌은 반드시 별도로 정리할 것.
3. 출처가 불분명한 통계, 수치, 기관명, 연구명은 지어내지 말 것.
4. 핵심 주장에는 가능한 한 실제 자료, 보고서, 논문, 공공기관 자료를 반영할 것.
5. 단순 나열이 아니라 주장 → 근거 → 해석 흐름으로 쓸 것.
6. body에는 오직 본문만 넣을 것.
7. body 안에 "참고문헌", "참고자료", "인용" 섹션 제목을 따로 만들지 말 것.
8. references에는 참고문헌만 배열로 넣을 것.
9. citations 필드는 항상 빈 배열로 반환해도 된다.
10. JSON 바깥에 어떤 설명도 출력하지 말 것.
11. {get_length_rule(payload["length"])}
12. 문체는 {payload["tone"]}를 기본으로 할 것.
13. {get_reference_style_instruction(reference_style)}

{get_mode_rules(payload["mode"])}

{get_assignment_rules(payload)}

{get_template_structure(payload)}

추가 규칙:
- 본문 안의 인용은 자연스럽게 넣을 것.
- 참고문헌은 references 배열에만 넣을 것.
- citations는 빈 배열이어도 괜찮다.
- output_language는 반드시 "ko"로 넣을 것.

반드시 아래 JSON 형식으로만 응답:
{{
  "title": "문서 제목",
  "body": "문서 본문 전체",
  "references": [
    "참고문헌 1",
    "참고문헌 2"
  ],
  "citations": [],
  "output_language": "ko"
}}
"""


def build_refine_prompt(action: str, title: str, body: str, template: str, tone: str) -> str:
    action_map = {
        "polish_style": """
목표:
- 문체를 더 매끄럽고 자연스럽게 다듬을 것.
- 의미는 유지하되 표현과 문단 흐름을 개선할 것.
""",
        "strengthen_evidence": """
목표:
- 기존 본문을 바탕으로 근거와 사례를 더 보강한 새로운 문서를 작성할 것.
- 본문 안에 인용은 자연스럽게 포함할 수 있다.
- 참고문헌은 별도로 정리할 것.
""",
        "expand_conclusion": """
목표:
- 결론을 더 풍부하게 확장한 새로운 문서를 작성할 것.
""",
        "presentation_summary": """
목표:
- 발표용으로 더 명확하고 전달력 있게 정리한 새로운 문서를 작성할 것.
""",
        "add_critical_view": """
목표:
- 비판적 관점, 한계, 반론을 보강한 새로운 문서를 작성할 것.
"""
    }

    selected_action = action_map.get(action, action_map["polish_style"])

    return f"""
너는 기존 문서를 후처리하여 새로운 결과 문서를 생성하는 편집 엔진이다.

문서 유형: {template}
기본 문체: {tone}

{selected_action}

규칙:
1. 결과는 원문을 덮어쓰는 수정본이 아니라, 새롭게 정리된 완성본처럼 출력할 것.
2. 본문 안에 참고문헌 섹션 제목을 따로 넣지 말 것.
3. 참고문헌은 references 배열로만 반환할 것.
4. JSON 형식으로만 출력할 것.
5. output_language는 "ko"로 넣을 것.

반드시 아래 JSON 형식으로만 응답:
{{
  "title": "새 결과 제목",
  "body": "새 결과 본문",
  "references": [
    "참고문헌 1"
  ],
  "citations": [],
  "output_language": "ko"
}}

원문 제목:
{title}

원문 본문:
{body}
"""


def extract_json_text(raw_text: str) -> str:
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


def cleanup_body_text(body: str) -> str:
    if not body:
        return body

    text = body.strip()
    patterns = [
        r'\n#\s*참고문헌[\s\S]*$',
        r'\n#\s*참고자료[\s\S]*$',
        r'\n#\s*인용[\s\S]*$',
        r'\n참고문헌[\s\S]*$',
        r'\n참고자료[\s\S]*$',
        r'\n인용[\s\S]*$'
    ]

    for pattern in patterns:
        text = re.sub(pattern, '', text, flags=re.IGNORECASE)

    return text.strip()


def cleanup_list_items(items: List[str]) -> List[str]:
    cleaned = []
    for item in items:
        value = normalize_value(item)
        if not value:
            continue
        if "파싱에 실패" in value:
            continue
        cleaned.append(value)
    return cleaned


def parse_model_response(raw_text: str) -> Tuple[str, str, List[str], List[str], str]:
    cleaned_text = extract_json_text(raw_text)

    try:
        parsed = json.loads(cleaned_text)

        title = normalize_value(parsed.get("title"), "생성된 문서")
        body = cleanup_body_text(normalize_value(parsed.get("body"), ""))
        references = parsed.get("references", [])
        citations = parsed.get("citations", [])
        output_language = normalize_value(parsed.get("output_language"), "ko")

        if not isinstance(references, list):
            references = [str(references)]

        if not isinstance(citations, list):
            citations = [str(citations)]

        references = cleanup_list_items(references)
        citations = cleanup_list_items(citations)

        if not title:
            title = "생성된 문서"

        if not body:
            body = "본문을 받아오지 못했습니다."

        return title, body, references, citations, output_language

    except Exception:
        title = "생성된 문서"
        body = cleanup_body_text(raw_text.strip())
        return title, body, [], [], "ko"


def safe_filename(title: str, ext: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "", title).strip()
    cleaned = cleaned.replace(" ", "_")
    if not cleaned:
        cleaned = f"document_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    return f"{cleaned}.{ext}"


def set_run_korean_font(run, font_name="Malgun Gothic", font_size=11, bold=False):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold

    r = run._element
    r.rPr.rFonts.set(qn("w:ascii"), font_name)
    r.rPr.rFonts.set(qn("w:hAnsi"), font_name)
    r.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    r.rPr.rFonts.set(qn("w:cs"), font_name)


def create_docx_file(title: str, body: str, references: List[str]) -> BytesIO:
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Malgun Gothic"
    style.font.size = Pt(11)
    style._element.rPr.rFonts.set(qn("w:ascii"), "Malgun Gothic")
    style._element.rPr.rFonts.set(qn("w:hAnsi"), "Malgun Gothic")
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
    style._element.rPr.rFonts.set(qn("w:cs"), "Malgun Gothic")

    title_para = doc.add_paragraph()
    title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    title_run = title_para.add_run(title)
    set_run_korean_font(title_run, font_name="Malgun Gothic", font_size=16, bold=True)

    doc.add_paragraph("")

    for line in body.split("\n"):
        stripped = line.strip()

        if not stripped:
            doc.add_paragraph("")
            continue

        if stripped.startswith("# "):
            p = doc.add_paragraph()
            r = p.add_run(stripped[2:].strip())
            set_run_korean_font(r, font_name="Malgun Gothic", font_size=14, bold=True)
        else:
            p = doc.add_paragraph()
            r = p.add_run(stripped)
            set_run_korean_font(r, font_name="Malgun Gothic", font_size=11, bold=False)

    if references:
        doc.add_paragraph("")
        p = doc.add_paragraph()
        r = p.add_run("참고문헌")
        set_run_korean_font(r, font_name="Malgun Gothic", font_size=13, bold=True)

        for ref in references:
            p = doc.add_paragraph(style="List Bullet")
            r = p.add_run(ref)
            set_run_korean_font(r, font_name="Malgun Gothic", font_size=11, bold=False)

    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


def register_korean_font() -> str:
    font_path = os.path.join(app.root_path, "fonts", "NanumGothic.ttf")
    font_name = "Helvetica"

    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont("NanumGothic", font_path))
        font_name = "NanumGothic"

    return font_name


def create_pdf_file(title: str, body: str, references: List[str]) -> BytesIO:
    file_stream = BytesIO()
    font_name = register_korean_font()

    doc = SimpleDocTemplate(
        file_stream,
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "KoreanTitle",
        parent=styles["Title"],
        fontName=font_name,
        fontSize=16,
        leading=22,
        spaceAfter=18
    )

    heading_style = ParagraphStyle(
        "KoreanHeading",
        parent=styles["Heading2"],
        fontName=font_name,
        fontSize=13,
        leading=18,
        spaceBefore=12,
        spaceAfter=8
    )

    body_style = ParagraphStyle(
        "KoreanBody",
        parent=styles["BodyText"],
        fontName=font_name,
        fontSize=10.5,
        leading=16,
        spaceAfter=8
    )

    story = []
    story.append(Paragraph(title.replace("\n", "<br/>"), title_style))
    story.append(Spacer(1, 10))

    for line in body.split("\n"):
        stripped = line.strip()
        if not stripped:
            story.append(Spacer(1, 6))
            continue

        if stripped.startswith("# "):
            story.append(Paragraph(stripped[2:].strip(), heading_style))
        else:
            safe_line = stripped.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(safe_line, body_style))

    if references:
        story.append(Spacer(1, 12))
        story.append(Paragraph("참고문헌", heading_style))
        for ref in references:
            safe_ref = f"• {ref}".replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(safe_ref, body_style))

    doc.build(story)
    file_stream.seek(0)
    return file_stream


@app.errorhandler(Exception)
def handle_exception(e):
    return jsonify({
        "success": False,
        "message": f"서버 내부 오류: {str(e)}"
    }), 500


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "success": True,
        "status": "ok",
        "service": "ai-document-generator",
        "model": MODEL_NAME
    })


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(silent=True) or {}
    payload = sanitize_input(data)

    prompt = build_prompt(payload)

    response = client.responses.create(
        model=MODEL_NAME,
        tools=[{"type": "web_search_preview"}],
        input=prompt
    )

    raw_text = response.output_text.strip()
    title, body, references, citations, output_language = parse_model_response(raw_text)

    return jsonify({
        "success": True,
        "title": title,
        "body": body,
        "references": references,
        "citations": citations,
        "output_language": output_language
    })


@app.route("/refine", methods=["POST"])
def refine():
    data = request.get_json(silent=True) or {}

    action = normalize_value(data.get("action"), "polish_style")
    title = normalize_value(data.get("title"), "생성된 문서")
    body = normalize_value(data.get("body"), "")
    template = normalize_value(data.get("template"), "report")
    tone = normalize_value(data.get("tone"), "자연스럽고 이해하기 쉽게")

    prompt = build_refine_prompt(action, title, body, template, tone)

    response = client.responses.create(
        model=MODEL_NAME,
        tools=[{"type": "web_search_preview"}],
        input=prompt
    )

    raw_text = response.output_text.strip()
    new_title, new_body, references, citations, output_language = parse_model_response(raw_text)

    return jsonify({
        "success": True,
        "title": new_title,
        "body": new_body,
        "references": references,
        "citations": citations,
        "output_language": output_language
    })


@app.route("/download/docx", methods=["POST"])
def download_docx():
    data = request.get_json(silent=True) or {}

    title = normalize_value(data.get("title"), "생성된 문서")
    body = normalize_value(data.get("body"), "")
    references = data.get("references", [])

    if not isinstance(references, list):
        references = []

    file_stream = create_docx_file(title, body, references)
    filename = safe_filename(title, "docx")

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.route("/download/pdf", methods=["POST"])
def download_pdf():
    data = request.get_json(silent=True) or {}

    title = normalize_value(data.get("title"), "생성된 문서")
    body = normalize_value(data.get("body"), "")
    references = data.get("references", [])

    if not isinstance(references, list):
        references = []

    file_stream = create_pdf_file(title, body, references)
    filename = safe_filename(title, "pdf")

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/pdf"
    )


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)