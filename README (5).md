# AI Exam Generator

Turns a lecture document (PDF, Word file, or a plain photo/scan) into a ready-to-print exam paper and a matching answer key — for any subject, not just one.

## What it does

1. **Reads your material.** Upload a `.pdf`, `.docx`, or an image (`.png`/`.jpg`/`.jpeg`/`.bmp`/`.tiff`/`.webp`). Typed text is extracted directly; scanned pages, handwriting, and embedded diagrams are read with a vision model instead of plain OCR, so content that's only "shown in a diagram" still gets captured.
2. **Drafts questions from it** — nothing invented, only what the material actually supports:
   - **Section 1 (Objective)** — at least 4 options per question, exactly one correct.
   - **Section 2 (Theory)** — questions that can expand into lettered sub-questions when the topic is complex enough.
   - You choose how many of each.
3. **Formats it like a real exam**, A4 size, with blank fields up top for School, Term, Class, Date, Time, and Subject, each section with its own instructions line.
4. **Writes a matching answer key**, including a model answer for every theory question.
5. **Works across subjects** — math symbols (`√`, fractions, exponents, Greek letters, set notation, etc.) and simple data tables render as real formatted content in the Word file, not raw text or LaTeX code.

## Tech stack

- **[Streamlit](https://streamlit.io)** — the web interface.
- **[Groq](https://groq.com)** — one model reads vision content (scans, handwriting, diagrams), a separate model drafts the questions.
- **python-docx** — builds the `.docx` exam and answer key.
- **PyMuPDF, pytesseract, OpenCV, poppler** — document/image ingestion.

## Setup

### 1. Requirements

This repo expects two files at its root:

- `requirements.txt` — Python packages (Streamlit, groq, python-docx, PyMuPDF, pytesseract, opencv-python-headless, pillow).
- `packages.txt` — system packages Streamlit Cloud needs installed alongside Python (`tesseract-ocr`, `poppler-utils`, `libgl1`, `libglib2.0-0`).

### 2. API key

Get a free key at [console.groq.com](https://console.groq.com). **Never commit it to the repo.**

- **Streamlit Community Cloud:** app dashboard → Settings → Secrets, then add:
  ```toml
  GROQ_API_KEY = "your-actual-key-here"
  ```
- **Running locally:** create `.streamlit/secrets.toml` with the same content, and make sure that path is in `.gitignore`.

### 3. Deploy

Push `app.py`, `requirements.txt`, and `packages.txt` to a GitHub repo, then connect it at [share.streamlit.io](https://share.streamlit.io). Streamlit Cloud rebuilds automatically on every push to the connected branch.

## Using it

1. Open the app and upload your lecture material.
2. Fill in the exam details (School, Term, Class, Date, Time, Subject) — leave any blank to fill in by hand later.
3. Set each section's instructions and how many objective/theory questions you want.
4. Click **Generate exam**.
5. Download `exam.docx` and `answer_key.docx`.

## Known limitations (current version)

- A large request (hundreds of questions) can take tens of minutes — keep the browser tab open while it runs.
- If a document genuinely doesn't contain enough distinct content to hit a very large question count, the app reports fewer questions than requested rather than repeating itself or inventing ungrounded ones.
- Diagrams from the source material are described in text to ground questions, but the diagram images themselves aren't yet embedded into the generated exam.
- The on-disk cache is temporary per run — Streamlit Cloud's storage doesn't persist across restarts/redeploys the way Google Drive did in the original Colab version.
