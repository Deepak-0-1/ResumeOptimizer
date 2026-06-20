# 🚀 Resume Optimizer v2

**ATS & Human-Friendly Resume Tailoring Tool** — Upload your resume + job description → get an optimized, ATS-ready resume as a DOCX download.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-3.0+-green?logo=flask&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow)

---

## ✨ Features

- **Upload or Paste** — Supports `.docx`, `.pdf`, and `.txt` resume uploads, or direct paste
- **AI-Powered Optimization** — Uses LLM to rewrite your resume targeting the specific job description
- **ATS Keyword Matching** — Mirrors exact keywords and phrases from the JD
- **DOCX Output** — Downloads a professionally formatted `.docx` using your template
- **Streaming Analysis** — Real-time optimization analysis with ATS compatibility score
- **No Data Stored** — Your resume and JD are processed in-memory only

## 📸 How It Works

1. **Paste/Upload** your resume and the target job description
2. Click **Optimize Resume**
3. The AI rewrites your resume to match the JD (without fabricating anything)
4. **Download** the optimized resume as a `.docx`
5. Review the **analysis** with keyword matching and ATS score

---

## 🛠️ Setup

### Prerequisites

- [Python 3.10+](https://www.python.org/downloads/)
- [uv](https://docs.astral.sh/uv/getting-started/installation/) (Python package manager)

### Installation

```bash
# Clone the repository
git clone https://github.com/Deepak-0-1/ResumeOptimizer.git
cd ResumeOptimizer

# Install dependencies with uv
uv sync

# Create your .env file with your API key
echo AGENTROUTER_API_KEY=your_api_key_here > .env
```

### Running

```bash
uv run python app.py
```

Then open **http://localhost:5000** in your browser.

---

## 📁 Project Structure

```
resume_optimizer/
├── app.py                          # Flask app — routes & LLM integration
├── extractors.py                   # Text extraction from DOCX/PDF/TXT
├── docx_builder.py                 # DOCX assembly with template styling
├── Technical_Resume_Template.docx  # Resume template (formatting reference)
├── templates/
│   └── index.html                  # Frontend UI (single-page app)
├── pyproject.toml                  # Project config & dependencies
├── uv.lock                        # Locked dependency versions
├── requirements.txt                # Legacy pip requirements
├── .env                            # API key (not committed)
└── .gitignore
```

## 🔧 Configuration

| Variable | Description |
|---|---|
| `AGENTROUTER_API_KEY` | Your API key for the LLM provider (AgentRouter) |

Create a `.env` file in the project root:

```env
AGENTROUTER_API_KEY=sk-your-key-here
```

## 📦 Dependencies

| Package | Purpose |
|---|---|
| `flask` | Web framework |
| `openai` | LLM API client |
| `python-docx` | DOCX reading & writing |
| `docxtpl` | DOCX template engine |
| `pdfplumber` | PDF text extraction |
| `python-dotenv` | Environment variable management |

---

## ⚠️ Important Notes

- The AI **never fabricates** experience or skills — it only reorganizes and rephrases existing content
- Your resume template (`Technical_Resume_Template.docx`) defines the output formatting
- All processing happens in-memory; no data is persisted

## 📄 License

MIT License — feel free to use and modify.
