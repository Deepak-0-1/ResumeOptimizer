"""
Resume Optimizer v2 — ATS & Human-Friendly Resume Tailoring Tool
Upload resume (DOCX/PDF/paste) + JD → LLM optimizes content → DOCX output
"""

import sys
import os
import json
from dotenv import load_dotenv

load_dotenv()
from flask import Flask, request, jsonify, render_template, Response, send_file
from openai import OpenAI
from extractors import extract_text
from docx_builder import build_resume, parse_llm_json

# Ensure console handles Unicode characters on Windows
sys.stdout.reconfigure(encoding='utf-8')

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB upload limit

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "Technical_Resume_Template.docx")


# ─────────────────────────────────────────────
# AgentRouter LLM Client (same pattern as testapp2.py)
# ─────────────────────────────────────────────
class AgentRouterClient:
    def __init__(self, api_key: str):
        self.client = OpenAI(
            base_url="https://agentrouter.org/v1",
            api_key=api_key,
            default_headers={
                "Originator": "codex_cli_rs",
                "User-Agent": "codex_cli_rs/0.101.0 (Mac OS 26.0.1; arm64) Apple_Terminal/464",
                "Version": "0.101.0"
            }
        )

    def ask(self, prompt: str, system_prompt: str = "", model: str = "deepseek-v4-flash") -> str:
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})
        try:
            response = self.client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=0.4,
            )
            return response.choices[0].message.content
        except Exception as e:
            return f"Error: {e}"

    def ask_stream(self, prompt: str, system_prompt: str = "", model: str = "deepseek-v4-flash"):
        """Generator that yields content chunks for streaming responses."""
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})
        try:
            stream = self.client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=0.4,
                stream=True,
            )
            for chunk in stream:
                if chunk.choices and chunk.choices[0].delta and chunk.choices[0].delta.content:
                    yield chunk.choices[0].delta.content
        except Exception as e:
            yield f"\n\n**Error:** {e}"


# ─────────────────────────────────────────────
# Prompt Engineering — Two-Phase Approach
# ─────────────────────────────────────────────

# Phase 1: Structured JSON output for DOCX assembly
SYSTEM_PROMPT_JSON = """You are an elite resume optimization expert. Your job is to optimize a resume for a specific job description to maximize both ATS score AND human recruiter appeal.

STRICT RULES:
1. NEVER fabricate experience, skills, or achievements. Only reorganize, rephrase, and emphasize what already exists.
2. If the candidate lacks a required skill, do NOT add it. Instead, highlight transferable skills.
3. Use the EXACT keywords and phrases from the JD wherever truthfully applicable.
4. Mirror the job title language from the JD.
5. Front-load each bullet point with strong action verbs.
6. Quantify achievements wherever possible (numbers, percentages, dollar amounts).
7. Ensure the Professional Summary is a targeted 3-4 line pitch directly addressing the JD requirements.
8. Keep ALL existing sections from the resume. Do not drop any section.

You MUST respond with ONLY a valid JSON object (no markdown, no explanation, no code fences). The JSON must follow this EXACT structure:

{
  "name": "Full Name",
  "subtitle": "Role 1  |  Role 2  |  Role 3",
  "contact_line": "City, Country  •  email@example.com  •  +XX-XXXXX-XXXXX  •  linkedin.com/in/handle  •  github.com/handle",
  "summary": "3-4 sentence professional summary tailored to the JD...",
  "skills": [
    {"category": "Languages", "items": "Python, SQL, Bash"},
    {"category": "Data Eng", "items": "PySpark, Kafka, Airflow, ..."},
    {"category": "Cloud / Infra", "items": "AWS (S3, Glue, ...), Docker, ..."},
    {"category": "AI / ML", "items": "LangChain, scikit-learn, ..."},
    {"category": "GenAI / LLMOps", "items": "RAG pipelines, Prompt Engineering, ..."},
    {"category": "Monitoring", "items": "MLflow, Evidently AI, ..."}
  ],
  "experience": [
    {
      "title": "Job Title",
      "dates": "Month Year – Present",
      "company": "Company Name",
      "location": "Remote / City, Country",
      "bullets": [
        "Strong action verb + quantified achievement + impact...",
        "Another bullet point..."
      ]
    }
  ],
  "projects": [
    {
      "name": "Project Name",
      "tech": "Tech1 · Tech2 · Tech3",
      "bullets": [
        "What you built + quantified impact...",
        "Another bullet..."
      ]
    }
  ],
  "education": [
    {
      "degree": "B.Tech in Computer Science",
      "institution": "University Name, City",
      "dates": "20XX – 20XX",
      "coursework": "Relevant Coursework: Subject1, Subject2, ..."
    }
  ],
  "certifications": [
    "Certification Name 1",
    "Certification Name 2"
  ],
  "additional": [
    "Additional info item 1",
    "Additional info item 2"
  ]
}"""

USER_PROMPT_JSON = """## CANDIDATE'S CURRENT RESUME:
{resume}

---

## TARGET JOB DESCRIPTION:
{jd}

---

Optimize this resume for the above job description. Return ONLY the JSON object as specified. No markdown, no code fences, no explanation."""


# Phase 2: Human-readable analysis (streamed to browser)
SYSTEM_PROMPT_ANALYSIS = """You are an elite resume optimization expert. You have already optimized a resume. Now provide a brief analysis. Be concise and actionable."""

USER_PROMPT_ANALYSIS = """I just optimized a resume for this job description. Provide a brief analysis:

## JOB DESCRIPTION:
{jd}

## OPTIMIZED RESUME CONTENT:
{optimized_summary}

---

Provide ONLY these sections (keep it concise):

### KEY CHANGES MADE
Bullet-point list of the major changes and WHY each improves ATS/recruiter performance.

### KEYWORD MATCH ANALYSIS
- Critical keywords from the JD
- Mark present ✅ or missing ❌

### ATS COMPATIBILITY SCORE
Rate 1-10 with brief justification.

### TIPS FOR THE CANDIDATE
2-3 actionable suggestions beyond the resume."""


# ─────────────────────────────────────────────
# API Key — loaded from .env
# ─────────────────────────────────────────────
API_KEY = os.environ.get("AGENTROUTER_API_KEY", "")
ai = AgentRouterClient(API_KEY)


# ─────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/extract", methods=["POST"])
def extract():
    """Extract text from an uploaded file (DOCX/PDF/TXT).
    Returns the extracted text for preview.
    """
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "Empty filename"}), 400

    try:
        file_bytes = file.read()
        text = extract_text(file_bytes, file.filename)
        return jsonify({"text": text, "filename": file.filename})
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Failed to extract text: {e}"}), 500


@app.route("/optimize", methods=["POST"])
def optimize():
    """Two-phase optimization:
    Phase 1: Get structured JSON from LLM (non-streamed, for DOCX)
    Phase 2: Stream analysis to browser (SSE)
    
    Returns SSE stream with:
    - First event: {"type": "json", "data": {...}} with the structured resume
    - Subsequent events: {"type": "stream", "content": "..."} with analysis text
    - Final event: [DONE]
    """
    data = request.get_json()
    resume = data.get("resume", "").strip()
    jd = data.get("jd", "").strip()

    if not resume or not jd:
        return jsonify({"error": "Both resume and job description are required."}), 400

    def generate():
        # ── Phase 1: Structured JSON optimization ──
        yield f"data: {json.dumps({'type': 'status', 'content': 'Optimizing resume content...'})}\n\n"

        json_prompt = USER_PROMPT_JSON.format(resume=resume, jd=jd)
        raw_response = ai.ask(json_prompt, system_prompt=SYSTEM_PROMPT_JSON)

        try:
            optimized_data = parse_llm_json(raw_response)
            # Send the structured JSON to the client
            yield f"data: {json.dumps({'type': 'json', 'data': optimized_data})}\n\n"
        except (json.JSONDecodeError, ValueError) as e:
            yield f"data: {json.dumps({'type': 'error', 'content': f'Failed to parse LLM response: {e}. Raw: {raw_response[:500]}'})}\n\n"
            yield "data: [DONE]\n\n"
            return

        # ── Phase 2: Stream analysis ──
        yield f"data: {json.dumps({'type': 'status', 'content': 'Generating analysis...'})}\n\n"

        # Create a summary of the optimized content for the analysis prompt
        optimized_summary = f"Name: {optimized_data.get('name', '')}\n"
        optimized_summary += f"Summary: {optimized_data.get('summary', '')}\n"
        optimized_summary += f"Skills: {json.dumps(optimized_data.get('skills', []))}\n"

        analysis_prompt = USER_PROMPT_ANALYSIS.format(
            jd=jd,
            optimized_summary=optimized_summary
        )

        for chunk in ai.ask_stream(analysis_prompt, system_prompt=SYSTEM_PROMPT_ANALYSIS):
            yield f"data: {json.dumps({'type': 'stream', 'content': chunk})}\n\n"

        yield "data: [DONE]\n\n"

    return Response(generate(), mimetype="text/event-stream")


@app.route("/download-docx", methods=["POST"])
def download_docx():
    """Generate and download the optimized resume as DOCX.
    
    Expects JSON body with the optimized_data dict.
    Uses the user's template to assemble the final document.
    """
    data = request.get_json()
    optimized_data = data.get("optimized_data")

    if not optimized_data:
        return jsonify({"error": "No optimized data provided"}), 400

    if not os.path.exists(TEMPLATE_PATH):
        return jsonify({"error": "Resume template not found. Please place Technical_Resume_Template.docx in the app directory."}), 404

    try:
        docx_buffer = build_resume(TEMPLATE_PATH, optimized_data)
        return send_file(
            docx_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='optimized_resume.docx'
        )
    except Exception as e:
        return jsonify({"error": f"Failed to generate DOCX: {e}"}), 500


# ─────────────────────────────────────────────
# Run
# ─────────────────────────────────────────────
if __name__ == "__main__":
    print("\n🚀 Resume Optimizer v2 running at http://localhost:5000\n")
    app.run(debug=True, host="0.0.0.0", port=5000)
