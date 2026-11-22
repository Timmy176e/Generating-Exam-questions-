"""
AI Exam Generator — Streamlit app.

Reads a lecture document (PDF, DOCX, or image), drafts an exam (objective +
theory sections) grounded strictly in that material, and renders both an
exam paper and a matching answer key as .docx files, in A4 format with
real math symbols/tables instead of raw LaTeX text.
"""

import os
import re
import io
import json
import time
import base64
import hashlib
import zipfile
import subprocess
import tempfile

import numpy as np
import cv2
import pytesseract
import fitz  # PyMuPDF
from PIL import Image
from docx import Document as DocxDocument
from docx import Document
from docx.shared import Cm, Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from groq import Groq
import streamlit as st

st.set_page_config(page_title="AI Exam Generator", page_icon="📝")

# ============================================================
# API key + client (the app's own shared Groq key, from Streamlit secrets)
# ============================================================

try:
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except Exception:
    GROQ_API_KEY = None

client = Groq(api_key=GROQ_API_KEY) if GROQ_API_KEY else None

VISION_MODEL = "qwen/qwen3.8-27b"
GENERATION_MODEL = "openai/gpt-oss-120b"

# IMPORTANT: this must be created ONCE per browser session, not on every
# script rerun (Streamlit reruns the whole file on every interaction,
# including a download-button click). st.session_state is what survives
# across reruns within one session; tempfile.mkdtemp() alone does not.
if "cache_dir" not in st.session_state:
    st.session_state.cache_dir = tempfile.mkdtemp()
CACHE_DIR = st.session_state.cache_dir


# ============================================================
# Ingestion: PDF / DOCX / image -> plain text (vision-assisted where needed)
# ============================================================

TABLE_DIAGRAM_HINTS = re.compile(
    r"\b(figure|diagram|table|shown below|chart|graph)\b", re.IGNORECASE
)


def preprocess_for_ocr(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    thresh = cv2.adaptiveThreshold(
        blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2
    )
    return thresh


def ocr_image(pil_image):
    img = cv2.cvtColor(np.array(pil_image.convert("RGB")), cv2.COLOR_RGB2BGR)
    processed = preprocess_for_ocr(img)
    return pytesseract.image_to_string(processed)


def page_needs_vision(page_text):
    if len(page_text.strip()) < 20:
        return True
    return bool(TABLE_DIAGRAM_HINTS.search(page_text))


def vision_read_image_bytes(img_bytes, mime_type="image/png"):
    if client is None:
        return None
    b64 = base64.b64encode(img_bytes).decode("utf-8")
    completion = client.chat.completions.create(
        model=VISION_MODEL,
        max_tokens=1000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": (
                    "Transcribe everything on this page, concisely but completely: all text, "
                    "equations (write them out in words or simple notation), tables "
                    "(reproduce as a markdown table, preserving every row and column), "
                    "and diagrams (describe briefly what each one shows). Keep the original order."
                )},
                {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{b64}"}},
            ],
        }],
    )
    return completion.choices[0].message.content


def get_page_text_layout(pdf_path, page_num):
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", "-f", str(page_num), "-l", str(page_num), pdf_path, "-"],
            capture_output=True, text=True, timeout=30
        )
        return result.stdout
    except Exception:
        return ""


def cache_path_for(path, page_num):
    base = os.path.basename(path).rsplit(".", 1)[0]
    return os.path.join(CACHE_DIR, f"{base}_page{page_num:03d}.txt")


def extract_text_from_pdf(path, debug=True, progress_callback=None):
    doc = fitz.open(path)
    total_pages = len(doc)
    full_text = ""
    for i, page in enumerate(doc):
        page_num = i + 1
        cache_file = cache_path_for(path, page_num)

        if os.path.exists(cache_file):
            with open(cache_file, "r", encoding="utf-8") as f:
                cached = f.read()
            if debug:
                print(f"Page {page_num}: using cached result ({len(cached)} chars)")
            full_text += cached + "\n"
            if progress_callback:
                progress_callback(page_num, total_pages)
            continue

        page_text = page.get_text()

        if page_needs_vision(page_text):
            if debug:
                print(f"Page {page_num}: text layer too thin/hints at a table or diagram "
                      f"({len(page_text.strip())} chars) -> trying vision")
            pix = page.get_pixmap(dpi=250)
            try:
                result = vision_read_image_bytes(pix.tobytes("png"))
            except Exception as e:
                print(f"  vision call raised an error: {e}")
                result = None

            if result:
                if debug:
                    print(f"  vision succeeded ({len(result)} chars)")
                with open(cache_file, "w", encoding="utf-8") as f:
                    f.write(result)
                full_text += result + "\n"
                if progress_callback:
                    progress_callback(page_num, total_pages)
                continue

            if client is None:
                print(f"  Page {page_num}: GROQ_API_KEY not set -> falling back to pytesseract OCR (lower quality)")
            else:
                print(f"  Page {page_num}: vision unavailable right now -> falling back to pytesseract OCR")
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            full_text += ocr_image(img) + "\n"
            # deliberately NOT cached — so we retry vision on this page once quota resets
        else:
            if debug:
                print(f"Page {page_num}: has a real text layer ({len(page_text.strip())} chars) -> layout extraction")
            layout_text = get_page_text_layout(path, page_num)
            result_text = layout_text if layout_text.strip() else page_text
            with open(cache_file, "w", encoding="utf-8") as f:
                f.write(result_text)
            full_text += result_text + "\n"

        if progress_callback:
            progress_callback(page_num, total_pages)
    return full_text


def _natural_key(name):
    numbers = re.findall(r'\d+', name)
    return int(numbers[-1]) if numbers else 0


def extract_docx_images(docx_path):
    """Returns (filename, image_bytes) for every embedded image, in document order."""
    images = []
    with zipfile.ZipFile(docx_path) as z:
        media_files = sorted(
            [n for n in z.namelist() if n.startswith("word/media/")],
            key=_natural_key
        )
        for name in media_files:
            images.append((name, z.read(name)))
    return images


def _image_mime_type(filename):
    ext = filename.lower().rsplit(".", 1)[-1]
    return {
        "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "gif": "image/gif", "bmp": "image/bmp",
    }.get(ext)  # returns None for formats we can't send to vision (e.g. .emf/.wmf drawings)


def is_image_worth_describing(img_bytes, min_width=100, min_height=100):
    try:
        img = Image.open(io.BytesIO(img_bytes))
        w, h = img.size
        return w >= min_width and h >= min_height
    except Exception:
        return True  # can't inspect it -> don't silently skip, let vision try


def extract_text_from_docx(path, debug=True, progress_callback=None):
    doc = DocxDocument(path)
    text = "\n".join(p.text for p in doc.paragraphs)

    images = extract_docx_images(path)
    base = os.path.basename(path).rsplit(".", 1)[0]

    if not images:
        if progress_callback:
            progress_callback(1, 1)
        return text

    total = len(images)
    text += "\n\n--- DIAGRAMS AND FIGURES FROM THIS DOCUMENT ---\n"
    for i, (name, img_bytes) in enumerate(images, start=1):
        cache_file = os.path.join(CACHE_DIR, f"{base}_diagram{i:03d}.txt")

        if os.path.exists(cache_file):
            with open(cache_file, "r", encoding="utf-8") as f:
                cached = f.read()
            if debug:
                print(f"Diagram {i} ({name}): using cached result ({len(cached)} chars)")
            text += f"\n[Diagram {i}]: {cached}\n"
            if progress_callback:
                progress_callback(i, total)
            continue

        if not is_image_worth_describing(img_bytes):
            if debug:
                print(f"Diagram {i} ({name}): too small, likely an icon/decoration -> skipping")
            if progress_callback:
                progress_callback(i, total)
            continue

        mime_type = _image_mime_type(name)
        if mime_type is None:
            if debug:
                print(f"Diagram {i} ({name}): unsupported image format, skipping")
            if progress_callback:
                progress_callback(i, total)
            continue

        if debug:
            print(f"Diagram {i} ({name}): sending to vision...")
        try:
            description = vision_read_image_bytes(img_bytes, mime_type=mime_type)
        except Exception as e:
            print(f"  vision call raised an error: {e}")
            description = None

        if description:
            if debug:
                print(f"  described ({len(description)} chars)")
            with open(cache_file, "w", encoding="utf-8") as f:
                f.write(description)
            text += f"\n[Diagram {i}]: {description}\n"
        else:
            print("  could not describe right now (quota or key issue) -> will retry next run")

        if progress_callback:
            progress_callback(i, total)
    return text


def _guess_mime_type(path):
    ext = os.path.splitext(path)[1].lower()
    return {
        ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".bmp": "image/bmp", ".tiff": "image/tiff", ".webp": "image/webp",
    }.get(ext, "image/png")


def extract_text_from_image(path, debug=True, progress_callback=None):
    """For a photo/scan sent on its own (not inside a PDF or DOCX)."""
    if progress_callback:
        progress_callback(0, 1)

    with open(path, "rb") as f:
        img_bytes = f.read()

    cache_key = hashlib.md5(img_bytes).hexdigest()
    cache_file = os.path.join(CACHE_DIR, f"image_{cache_key}.txt")
    if os.path.exists(cache_file):
        if debug:
            print(f"Using cached result for {path}")
        with open(cache_file, "r") as f:
            text = f.read()
        if progress_callback:
            progress_callback(1, 1)
        return text

    if debug:
        print(f"Reading {path} with vision...")
    text = vision_read_image_bytes(img_bytes, mime_type=_guess_mime_type(path))

    if text:
        with open(cache_file, "w") as f:
            f.write(text)
        if progress_callback:
            progress_callback(1, 1)
        return text

    if debug:
        print("Vision unavailable -> falling back to pytesseract OCR")
    result = ocr_image(Image.open(path))
    if progress_callback:
        progress_callback(1, 1)
    return result


def extract_text(path, debug=True, progress_callback=None):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(path, debug=debug, progress_callback=progress_callback)
    elif ext == ".docx":
        return extract_text_from_docx(path, debug=debug, progress_callback=progress_callback)
    elif ext in (".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".webp"):
        return extract_text_from_image(path, debug=debug, progress_callback=progress_callback)
    else:
        raise ValueError(
            f"Unsupported file type '{ext}'. Supported: .pdf, .docx, .png, .jpg, .jpeg, .bmp, .tiff, .webp"
        )


# ============================================================
# Question generation (chunked, deduplicated, retrying, table-aware)
# ============================================================

QUESTION_GEN_PROMPT = """You are an exam-setting assistant. Based ONLY on the study material given below, draft exam questions.

Return STRICT JSON with this exact shape, and nothing else (no markdown, no commentary):

{{
  "section_1_objective": [
    {{
      "question": "...",
      "options": {{"A": "...", "B": "...", "C": "...", "D": "..."}},
      "correct_option": "A"
    }}
  ],
  "section_2_theory": [
    {{
      "question": "...",
      "sub_questions": ["...", "..."],
      "model_answer": "..."
    }}
  ]
}}

Rules:
- Section 1: exactly {num_objective} objective questions. Each has exactly 4 options, only one correct. Wrong options must be plausible, not silly or obviously wrong.
- Section 2: exactly {num_theory} theory questions. Add sub-questions only when a question's complexity genuinely calls for breaking it into parts; otherwise use an empty list.
- For each Section 2 question, also write a concise model_answer that correctly answers it (and its sub-questions, if any), based strictly on the material given.
- Sub-questions in section_2_theory must be plain question text only — do NOT prefix them with letters or numbers like "a)", "(a)", or "1.". The exam layout adds its own lettering automatically.
- Base every question strictly on the material given below. Do not invent facts it doesn't support, even if you believe you know the standard answer from general knowledge — use only what this specific material states.
- If the material references a diagram, figure, or image (e.g. "as shown below", "the figure shows") but does not describe its content in words anywhere in the material given, do NOT write a question or sub-question that depends on knowing what that diagram shows.
- Keep all mathematical notation and symbols EXACTLY as written in the material. For example, if the material writes a time shift as $t_o$, your question and model answer must also use $t_o$ — never substitute a different symbol such as $\\tau$ or $\\Delta t$, even if that symbol is more conventional in textbooks generally.
- Spread questions across the material rather than clustering on the first paragraph.

STUDY MATERIAL:
\"\"\"
{material}
\"\"\"
"""

_TABLE_FORMAT_INSTRUCTIONS = (
    "\n\nIf any question or its model answer needs to present data in a table "
    "(e.g. a truth table, a data table, a comparison table), format it as a "
    "Markdown-style pipe table exactly like this, with a header row, a "
    "separator row of dashes, and then the data rows:\n"
    "| Header 1 | Header 2 |\n"
    "|----------|----------|\n"
    "| value    | value    |\n"
    "Do not use any other table format, and do not describe a table in prose "
    "when the material itself provides one worth reproducing exactly."
)


def _format_avoid_block(avoid_questions, max_chars=2500):
    if not avoid_questions:
        return ""
    lines, total = [], 0
    for q in reversed(avoid_questions):
        line = f"- {q}"
        if total + len(line) > max_chars:
            break
        lines.append(line)
        total += len(line)
    lines.reverse()
    return (
        "\n\nQUESTIONS ALREADY ASKED FROM THIS MATERIAL — do not repeat these or ask "
        "close rewordings of them. If the material genuinely does not contain enough "
        "additional distinct, well-grounded content, return fewer questions than "
        "requested rather than duplicating or inventing ungrounded ones:\n"
        + "\n".join(lines)
    )


def _normalize_for_dedupe(s):
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def _extract_question_texts(result):
    texts = []
    for q in result.get("section_1_objective", []):
        if q.get("question"):
            texts.append(q["question"].strip())
    for q in result.get("section_2_theory", []):
        if q.get("question"):
            texts.append(q["question"].strip())
        for sub in (q.get("sub_questions") or []):
            if isinstance(sub, str) and sub.strip():
                texts.append(sub.strip())
    return texts


def _dedupe_questions(result):
    seen = set()
    removed = 0

    deduped_obj = []
    for q in result.get("section_1_objective", []):
        key = _normalize_for_dedupe(q.get("question", ""))
        if key and key not in seen:
            seen.add(key)
            deduped_obj.append(q)
        else:
            removed += 1

    deduped_theory = []
    for q in result.get("section_2_theory", []):
        key = _normalize_for_dedupe(q.get("question", ""))
        if key and key not in seen:
            seen.add(key)
            deduped_theory.append(q)
        else:
            removed += 1

    result["section_1_objective"] = deduped_obj
    result["section_2_theory"] = deduped_theory
    return result, removed


def generate_questions(material, num_objective=5, num_theory=3, avoid_questions=None):
    base_prompt = QUESTION_GEN_PROMPT.format(
        material=material, num_objective=num_objective, num_theory=num_theory
    )
    prompt = base_prompt + _TABLE_FORMAT_INSTRUCTIONS + _format_avoid_block(avoid_questions)

    completion = client.chat.completions.create(
        model=GENERATION_MODEL,
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"},
        temperature=0.4,
        max_tokens=3000,
    )
    raw = completion.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.strip("`")
        raw = raw.split("\n", 1)[1] if "\n" in raw else raw
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        print("Couldn't parse as JSON — raw model output was:")
        print(raw)
        raise


def _call_with_retry(chunk, take_obj, take_theory, avoid_so_far, max_retries=3):
    for attempt in range(max_retries):
        try:
            return generate_questions(chunk, num_objective=take_obj, num_theory=take_theory,
                                       avoid_questions=avoid_so_far)
        except Exception as e:
            msg = str(e).lower()
            if "rate_limit" in msg or "429" in msg or "tpm" in msg:
                wait = 30 * (attempt + 1)
                print(f"    -> hit a rate limit, waiting {wait}s before retrying "
                      f"(attempt {attempt+1}/{max_retries})")
                time.sleep(wait)
                continue
            if "json_validate_failed" in msg or "failed to generate json" in msg:
                print(f"    -> model produced malformed JSON, retrying "
                      f"(attempt {attempt+1}/{max_retries})")
                continue
            raise
    raise RuntimeError("Still failing after several retries.")


def _chunk_text(text, max_chars=8000):
    """Split text into chunks of at most max_chars, preferring to break
    at paragraph boundaries rather than mid-sentence."""
    paragraphs = text.split("\n\n")
    chunks = []
    current = ""
    for para in paragraphs:
        candidate = (current + "\n\n" + para) if current else para
        if len(candidate) <= max_chars:
            current = candidate
        else:
            if current:
                chunks.append(current)
            if len(para) <= max_chars:
                current = para
            else:
                for k in range(0, len(para), max_chars):
                    chunks.append(para[k:k + max_chars])
                current = ""
    if current:
        chunks.append(current)
    return chunks if chunks else [text]


def _split_counts(total, n_parts):
    """Split `total` into `n_parts` non-negative integers that sum to it,
    as evenly as possible."""
    if n_parts <= 0:
        return []
    base = total // n_parts
    remainder = total % n_parts
    return [base + (1 if i < remainder else 0) for i in range(n_parts)]


def _merge_question_results(results_list):
    merged = {"section_1_objective": [], "section_2_theory": []}
    for result in results_list:
        merged["section_1_objective"].extend(result.get("section_1_objective", []))
        merged["section_2_theory"].extend(result.get("section_2_theory", []))
    return merged


def generate_questions_chunked(full_text, num_objective, num_theory,
                                max_chars_per_chunk=8000, max_questions_per_call=6,
                                delay_between_calls=15, checkpoint_path=None,
                                progress_callback=None):
    chunks = _chunk_text(full_text, max_chars=max_chars_per_chunk)
    n = len(chunks)
    obj_counts = _split_counts(num_objective, n)
    theory_counts = _split_counts(num_theory, n)
    total_requested = max(num_objective + num_theory, 1)

    est_calls = 0
    for i in range(n):
        remaining = obj_counts[i] + theory_counts[i]
        if remaining > 0:
            est_calls += -(-remaining // max_questions_per_call)  # ceil division

    est_minutes = (est_calls * delay_between_calls) / 60
    print(f"Document split into {n} chunk(s).")
    print(f"Requesting {num_objective} objective + {num_theory} theory "
          f"-> about {est_calls} API call(s), roughly {est_minutes:.0f}+ minutes "
          f"of deliberate pacing alone (actual API response time adds more on top).")

    if progress_callback:
        progress_callback(0, total_requested)

    results = []
    asked_so_far = []
    obj_done = 0
    theory_done = 0
    first_call = True

    for i, chunk in enumerate(chunks):
        remaining_obj = obj_counts[i]
        remaining_theory = theory_counts[i]
        while remaining_obj > 0 or remaining_theory > 0:
            take_obj = min(remaining_obj, max_questions_per_call)
            take_theory = min(remaining_theory, max(max_questions_per_call - take_obj, 0))

            if not first_call:
                time.sleep(delay_between_calls)
            first_call = False

            print(f"  Chunk {i+1}/{n}: requesting {take_obj} objective + {take_theory} theory "
                  f"({len(asked_so_far)} questions banked so far)")
            try:
                batch = _call_with_retry(chunk, take_obj, take_theory, asked_so_far)
                got_obj = len(batch.get("section_1_objective", []))
                got_theory = len(batch.get("section_2_theory", []))
                print(f"    -> got {got_obj} objective + {got_theory} theory back")
                results.append(batch)
                asked_so_far.extend(_extract_question_texts(batch))
                obj_done += got_obj
                theory_done += got_theory
                if progress_callback:
                    progress_callback(min(obj_done + theory_done, total_requested), total_requested)
            except Exception as e:
                print(f"    -> call failed, skipping this batch: {e}")

            remaining_obj -= take_obj
            remaining_theory -= take_theory

            if checkpoint_path and results:
                try:
                    merged_so_far, _ = _dedupe_questions(_merge_question_results(results))
                    with open(checkpoint_path, "w") as f:
                        json.dump(merged_so_far, f)
                except Exception as e:
                    print(f"    (could not save checkpoint: {e})")

    merged, removed = _dedupe_questions(_merge_question_results(results))
    final_obj = len(merged.get("section_1_objective", []))
    final_theory = len(merged.get("section_2_theory", []))

    if progress_callback:
        progress_callback(total_requested, total_requested)

    print(f"\nDone. Requested {num_objective} objective + {num_theory} theory.")
    print(f"Got {final_obj} objective + {final_theory} theory unique questions "
          f"({removed} duplicate/near-duplicate dropped).")
    if final_obj < num_objective or final_theory < num_theory:
        print("That's short of what you asked for — the document ran out of genuinely "
              "new, distinct, well-grounded content before reaching your target, rather "
              "than the model repeating itself or inventing filler to hit the number.")

    return merged


# ============================================================
# LaTeX-to-Word math rendering
# ============================================================

_LATEX_SYMBOLS = {
    r"\times": "×", r"\cdot": "·", r"\div": "÷", r"\pm": "±", r"\mp": "∓",
    r"\leq": "≤", r"\le": "≤", r"\geq": "≥", r"\ge": "≥",
    r"\leqslant": "≤", r"\geqslant": "≥",
    r"\neq": "≠", r"\ne": "≠", r"\approx": "≈", r"\equiv": "≡",
    r"\sim": "∼", r"\propto": "∝",
    r"\in": "∈", r"\notin": "∉", r"\subseteq": "⊆", r"\subset": "⊂",
    r"\supseteq": "⊇", r"\supset": "⊃", r"\cup": "∪", r"\cap": "∩",
    r"\setminus": "∖",
    r"\emptyset": "∅", r"\varnothing": "∅", r"\forall": "∀", r"\exists": "∃",
    r"\lnot": "¬", r"\land": "∧", r"\lor": "∨", r"\wedge": "∧", r"\vee": "∨",
    r"\oplus": "⊕", r"\otimes": "⊗",
    r"\rightarrow": "→", r"\to": "→", r"\Rightarrow": "⇒",
    r"\leftarrow": "←", r"\Leftarrow": "⇐",
    r"\leftrightarrow": "↔", r"\Leftrightarrow": "⇔",
    r"\rightleftharpoons": "⇌",
    r"\sum": "∑", r"\int": "∫", r"\partial": "∂", r"\nabla": "∇",
    r"\infty": "∞", r"\therefore": "∴", r"\because": "∵",
    r"\angle": "∠", r"\parallel": "∥", r"\perp": "⊥", r"\triangle": "△",
    r"\hbar": "ℏ", r"\ell": "ℓ", r"\Re": "ℜ", r"\Im": "ℑ",
    r"\alpha": "α", r"\beta": "β", r"\gamma": "γ", r"\delta": "δ",
    r"\epsilon": "ε", r"\varepsilon": "ε", r"\zeta": "ζ", r"\eta": "η",
    r"\theta": "θ", r"\iota": "ι", r"\kappa": "κ", r"\lambda": "λ",
    r"\mu": "μ", r"\nu": "ν", r"\xi": "ξ", r"\pi": "π", r"\rho": "ρ",
    r"\sigma": "σ", r"\tau": "τ", r"\upsilon": "υ", r"\phi": "φ",
    r"\chi": "χ", r"\psi": "ψ", r"\omega": "ω",
    r"\Gamma": "Γ", r"\Delta": "Δ", r"\Theta": "Θ", r"\Lambda": "Λ",
    r"\Xi": "Ξ", r"\Pi": "Π", r"\Sigma": "Σ", r"\Upsilon": "Υ",
    r"\Phi": "Φ", r"\Psi": "Ψ", r"\Omega": "Ω",
    r"\degree": "°", r"\circ": "°", r"\ldots": "…", r"\cdots": "⋯",
    r"\{": "{", r"\}": "}", r"\,": " ", r"\;": " ", r"\!": "",
    r"\quad": "    ", r"\qquad": "        ",
}
_SYMBOL_KEYS_BY_LEN = sorted(_LATEX_SYMBOLS.keys(), key=len, reverse=True)

_COMBINING_MARKS = {
    "vec": "⃗", "hat": "̂", "bar": "̄",
    "overline": "̅", "dot": "̇", "ddot": "̈",
}


def _consume_brace(s, i):
    depth, j = 1, i + 1
    start = j
    while j < len(s) and depth > 0:
        if s[j] == "{":
            depth += 1
        elif s[j] == "}":
            depth -= 1
            if depth == 0:
                break
        j += 1
    return s[start:j], j + 1


def _consume_bracket(s, i):
    depth, j = 1, i + 1
    start = j
    while j < len(s) and depth > 0:
        if s[j] == "[":
            depth += 1
        elif s[j] == "]":
            depth -= 1
            if depth == 0:
                break
        j += 1
    return s[start:j], j + 1


def _render_math_runs(paragraph, latex, sup=False, sub=False, bold=False):
    i, n = 0, len(latex)
    while i < n:
        ch = latex[i]

        if ch == "$":
            close = latex.find("$", i + 1)
            if close != -1:
                _render_math_runs(paragraph, latex[i+1:close], sup, sub, bold)
                i = close + 1
            else:
                r = paragraph.add_run("$")
                r.font.superscript, r.font.subscript, r.bold = sup, sub, bold
                i += 1
            continue

        if ch == "\\":
            if (latex.startswith("\\(", i) or latex.startswith("\\)", i)
                    or latex.startswith("\\[", i) or latex.startswith("\\]", i)):
                i += 2
                continue

            matched = False
            for cmd in _SYMBOL_KEYS_BY_LEN:
                if latex.startswith(cmd, i):
                    r = paragraph.add_run(_LATEX_SYMBOLS[cmd])
                    r.font.superscript, r.font.subscript, r.bold = sup, sub, bold
                    i += len(cmd)
                    matched = True
                    break
            if matched:
                continue

            if (latex.startswith(r"\frac{", i) or latex.startswith(r"\dfrac{", i)
                    or latex.startswith(r"\tfrac{", i)):
                if latex.startswith(r"\dfrac{", i):
                    i += len(r"\dfrac")
                elif latex.startswith(r"\tfrac{", i):
                    i += len(r"\tfrac")
                else:
                    i += len(r"\frac")
                a, i = _consume_brace(latex, i)
                b, i = _consume_brace(latex, i)
                _render_math_runs(paragraph, a, sup, sub, bold)
                r = paragraph.add_run("⁄"); r.bold = bold
                _render_math_runs(paragraph, b, sup, sub, bold)
                continue

            if latex.startswith(r"\sqrt", i):
                j = i + len(r"\sqrt")
                index_text = None
                if j < n and latex[j] == "[":
                    index_text, j = _consume_bracket(latex, j)
                if j < n and latex[j] == "{":
                    arg, j = _consume_brace(latex, j)
                else:
                    arg, j = (latex[j], j + 1) if j < n else ("", j)
                if index_text:
                    _render_math_runs(paragraph, index_text, True, False, bold)
                r = paragraph.add_run("√("); r.bold = bold
                _render_math_runs(paragraph, arg, sup, sub, bold)
                r = paragraph.add_run(")"); r.bold = bold
                i = j
                continue

            combo_matched = False
            for cmd, mark in _COMBINING_MARKS.items():
                prefix = "\\" + cmd + "{"
                if latex.startswith(prefix, i):
                    i += len(cmd) + 1
                    inner, i = _consume_brace(latex, i)
                    marked = "".join(c + mark for c in inner)
                    r = paragraph.add_run(marked)
                    r.font.superscript, r.font.subscript, r.bold = sup, sub, bold
                    combo_matched = True
                    break
            if combo_matched:
                continue

            if latex.startswith(r"\left", i):
                i += len(r"\left")
                continue
            if latex.startswith(r"\right", i):
                i += len(r"\right")
                continue

            m = re.match(r"\\([a-zA-Z]+)", latex[i:])
            if m:
                r = paragraph.add_run(m.group(1))
                r.font.superscript, r.font.subscript, r.bold = sup, sub, bold
                i += m.end()
                continue
            i += 1
            continue

        if ch == "^":
            i += 1
            if i < n and latex[i] == "{":
                token, i = _consume_brace(latex, i)
            else:
                token, i = (latex[i], i + 1) if i < n else ("", i)
            _render_math_runs(paragraph, token, True, False, bold)
            continue

        if ch == "_":
            i += 1
            if i < n and latex[i] == "{":
                token, i = _consume_brace(latex, i)
            else:
                token, i = (latex[i], i + 1) if i < n else ("", i)
            _render_math_runs(paragraph, token, False, True, bold)
            continue

        if ch in "{}":
            r = paragraph.add_run(ch)
            r.font.superscript, r.font.subscript, r.bold = sup, sub, bold
            i += 1
            continue

        j = i
        while j < n and latex[j] not in "\\^_{}$":
            j += 1
        r = paragraph.add_run(latex[i:j])
        r.font.superscript, r.font.subscript, r.bold = sup, sub, bold
        i = j


_UNICODE_ESCAPE_RE = re.compile(r"\\u([0-9a-fA-F]{4})")


def add_formatted_text(paragraph, text, bold=False):
    """Add text to a docx paragraph, converting LaTeX math into real
    superscripts/subscripts/symbols — whether wrapped in $...$ / \\(...\\) /
    \\[...\\], or written bare with no wrapper at all."""
    text = _UNICODE_ESCAPE_RE.sub(lambda m: chr(int(m.group(1), 16)), text)
    _render_math_runs(paragraph, text, bold=bold)


# ============================================================
# A4 document rendering (exam + answer key), table-aware
# ============================================================

def _setup_a4_page(doc):
    section = doc.sections[0]
    section.page_height = Mm(297)
    section.page_width = Mm(210)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    normal = doc.styles['Normal']
    normal.font.size = Pt(11)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(3)


def _strip_leading_label(s):
    return re.sub(r'^\s*[\(\[]?[a-zA-Z0-9]{1,2}[\)\].:]\s*', '', s).strip()


def _field_display(value):
    return value if value else "_" * 25


def _add_title(doc, header_info):
    school = header_info.get("School", "").strip() or "EXAMINATION"
    school_title = doc.add_heading(school.upper(), level=1)
    school_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    term = header_info.get("Term", "").strip()
    if term:
        term_title = doc.add_heading(f"{term.upper()} EXAMINATION", level=2)
        term_title.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _add_header_fields(doc, header_info):
    p1 = doc.add_paragraph()
    p1.paragraph_format.tab_stops.add_tab_stop(Cm(9), WD_TAB_ALIGNMENT.LEFT)
    p1.add_run("Class: ").bold = True
    p1.add_run(_field_display(header_info.get("Class", "")))
    p1.add_run("\t")
    p1.add_run("Subject: ").bold = True
    p1.add_run(_field_display(header_info.get("Subject", "")))

    p2 = doc.add_paragraph()
    p2.paragraph_format.tab_stops.add_tab_stop(Cm(9), WD_TAB_ALIGNMENT.LEFT)
    p2.add_run("Date: ").bold = True
    p2.add_run(_field_display(header_info.get("Date", "")))
    p2.add_run("\t")
    p2.add_run("Time: ").bold = True
    p2.add_run(_field_display(header_info.get("Time", "")))
    p2.paragraph_format.space_after = Pt(10)


def _pack_options_line(options, letters, max_chars=95):
    lines, current, current_len = [], [], 0
    for letter in letters:
        text = f"{letter}. {options.get(letter, '')}"
        added_len = len(text) + (4 if current else 0)
        if current and current_len + added_len > max_chars:
            lines.append(current)
            current, current_len = [text], len(text)
        else:
            current.append(text)
            current_len += added_len
    if current:
        lines.append(current)
    return lines


def add_rich_content(doc, text, bold=False, space_after=None):
    """
    Render `text` into `doc` as paragraphs and/or real Word tables.
    Detects Markdown pipe tables (| a | b |) anywhere in the text and
    builds an actual table for them; everything else becomes ordinary
    paragraphs through add_formatted_text (so math still converts).
    """
    lines = text.split("\n")
    i, n = 0, len(lines)
    para_buffer = []
    state = {"last_para": None, "ended_in_table": False}

    def _is_table_row(line):
        s = line.strip()
        return s.startswith("|") and s.endswith("|") and s.count("|") >= 2

    def _is_separator_row(line):
        s = line.strip().strip("|")
        cells = [c.strip() for c in s.split("|")]
        return bool(cells) and all(re.match(r"^:?-{2,}:?$", c) for c in cells)

    def _flush_paragraph():
        if not para_buffer:
            return
        p = doc.add_paragraph()
        for idx, line in enumerate(para_buffer):
            if idx > 0:
                p.add_run().add_break()
            add_formatted_text(p, line, bold=bold)
        para_buffer.clear()
        state["last_para"] = p
        state["ended_in_table"] = False

    while i < n:
        line = lines[i]
        if (_is_table_row(line) and i + 1 < n and _is_table_row(lines[i + 1])
                and _is_separator_row(lines[i + 1])):
            _flush_paragraph()
            header_cells = [c.strip() for c in line.strip().strip("|").split("|")]
            j = i + 2
            data_rows = []
            while j < n and _is_table_row(lines[j]):
                row_cells = [c.strip() for c in lines[j].strip().strip("|").split("|")]
                data_rows.append(row_cells)
                j += 1
            ncols = len(header_cells)
            table = doc.add_table(rows=1, cols=ncols)
            table.style = "Table Grid"
            for c, cell_text in enumerate(header_cells):
                add_formatted_text(table.rows[0].cells[c].paragraphs[0], cell_text, bold=True)
            for row_cells in data_rows:
                row = table.add_row()
                for c in range(ncols):
                    val = row_cells[c] if c < len(row_cells) else ""
                    add_formatted_text(row.cells[c].paragraphs[0], val)
            i = j
            state["last_para"] = None
            state["ended_in_table"] = True
            continue

        if line.strip() == "":
            _flush_paragraph()
            i += 1
            continue

        para_buffer.append(line)
        i += 1

    _flush_paragraph()

    if space_after is not None:
        if state["ended_in_table"] or state["last_para"] is None:
            trailing = doc.add_paragraph()
            trailing.paragraph_format.space_after = space_after
        else:
            state["last_para"].paragraph_format.space_after = space_after


def render_exam(questions, header_info=None,
                 objective_instructions=None, theory_instructions=None,
                 output_path="exam.docx"):
    header_info = header_info or {}
    if objective_instructions is None:
        objective_instructions = "Answer ALL questions in this section."
    if theory_instructions is None:
        theory_instructions = "Answer ALL questions in this section."

    doc = Document()
    _setup_a4_page(doc)
    _add_title(doc, header_info)
    _add_header_fields(doc, header_info)

    doc.add_heading("SECTION 1: OBJECTIVE QUESTIONS", level=2)
    p1 = doc.add_paragraph()
    p1.add_run("INSTRUCTIONS: ").bold = True
    p1.add_run(objective_instructions)
    p1.paragraph_format.space_after = Pt(10)

    for i, q in enumerate(questions.get("section_1_objective", []), start=1):
        add_rich_content(doc, f"{i}. {q['question']}", bold=True)

        lines = _pack_options_line(q['options'], ["A", "B", "C", "D"])
        for idx, line_items in enumerate(lines):
            opt_para = doc.add_paragraph()
            add_formatted_text(opt_para, "    " + "     ".join(line_items))
            if idx == len(lines) - 1:
                opt_para.paragraph_format.space_after = Pt(10)

    doc.add_heading("SECTION 2: THEORY QUESTIONS", level=2)
    p2 = doc.add_paragraph()
    p2.add_run("INSTRUCTIONS: ").bold = True
    p2.add_run(theory_instructions)
    p2.paragraph_format.space_after = Pt(10)

    for i, q in enumerate(questions.get("section_2_theory", []), start=1):
        subs = q.get("sub_questions", [])
        add_rich_content(doc, f"{i}. {q['question']}", bold=True,
                          space_after=(Pt(10) if not subs else None))
        for j, sub in enumerate(subs, start=1):
            letter = chr(ord('a') + j - 1)
            sub_para = doc.add_paragraph()
            sub_para.add_run(f"    ({letter}) ")
            add_formatted_text(sub_para, _strip_leading_label(sub))
            if j == len(subs):
                sub_para.paragraph_format.space_after = Pt(10)

    doc.save(output_path)
    return output_path


def render_answer_key(questions, header_info=None, output_path="answer_key.docx"):
    header_info = header_info or {}
    doc = Document()
    _setup_a4_page(doc)
    _add_title(doc, header_info)

    key_label = doc.add_heading("ANSWER KEY", level=2)
    key_label.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("SECTION 1: OBJECTIVE ANSWERS", level=2)
    for i, q in enumerate(questions.get("section_1_objective", []), start=1):
        correct = q.get("correct_option", "?")
        option_text = q['options'].get(correct, '')
        ans_para = doc.add_paragraph()
        ans_para.add_run(f"{i}. {correct} — ")
        add_formatted_text(ans_para, option_text)

    doc.add_heading("SECTION 2: THEORY MODEL ANSWERS", level=2)
    for i, q in enumerate(questions.get("section_2_theory", []), start=1):
        add_rich_content(doc, f"{i}. {q['question']}", bold=True)
        add_rich_content(doc, q.get("model_answer", "(no model answer provided)"),
                          space_after=Pt(10))

    doc.save(output_path)
    return output_path


# ============================================================
# Streamlit UI
# ============================================================

st.title("AI Exam Generator")
st.caption("Upload lecture material, fill in the details, get an exam + answer key.")

if client is None:
    st.error(
        "No Groq API key is configured for this app. Add GROQ_API_KEY under "
        "Settings → Secrets in the Streamlit Cloud dashboard (or in "
        ".streamlit/secrets.toml when running locally), then reload."
    )

uploaded_file = st.file_uploader(
    "Upload a PDF, DOCX, or image of your lecture material",
    type=["pdf", "docx", "png", "jpg", "jpeg", "bmp", "tiff", "webp"],
)

with st.form("exam_details"):
    col1, col2 = st.columns(2)
    with col1:
        school = st.text_input("School")
        term = st.text_input("Term (e.g. First Term)")
        klass = st.text_input("Class")
    with col2:
        subject = st.text_input("Subject")
        date = st.text_input("Date")
        time_ = st.text_input("Time")

    objective_instructions = st.text_area(
        "Section 1 (Objective) instructions",
        value="Answer ALL questions in this section.",
    )
    theory_instructions = st.text_area(
        "Section 2 (Theory) instructions",
        value="Answer ALL questions in this section.",
    )

    col3, col4 = st.columns(2)
    with col3:
        num_objective = st.number_input("Section 1 (objective) questions", min_value=0, value=10, step=1)
    with col4:
        num_theory = st.number_input("Section 2 (theory) questions", min_value=0, value=3, step=1)

    submitted = st.form_submit_button("Generate exam")

if submitted:
    if client is None:
        st.error("Can't generate — no Groq API key is configured for this app (see message above).")
        st.stop()
    if uploaded_file is None:
        st.error("Please upload a file first.")
        st.stop()
    if num_objective == 0 and num_theory == 0:
        st.error("Enter at least one question in either section.")
        st.stop()

    header_info = {
        "School": school,
        "Term": term,
        "Class": klass,
        "Date": date,
        "Time": time_,
        "Subject": subject,
    }

    file_bytes = uploaded_file.getbuffer()
    file_hash = hashlib.md5(bytes(file_bytes)).hexdigest()
    temp_path = os.path.join(CACHE_DIR, uploaded_file.name)
    with open(temp_path, "wb") as f:
        f.write(file_bytes)

    try:
        # Skip re-extraction entirely if this exact file was already read
        # earlier in this same browser session.
        if st.session_state.get("last_file_hash") == file_hash and "last_extracted_text" in st.session_state:
            text = st.session_state["last_extracted_text"]
            st.info(f"Reusing text already extracted from this file earlier in this session "
                    f"({len(text)} characters) — skipping re-reading.")
        else:
            read_bar = st.progress(0, text="Reading your document... 0%")

            def _read_progress(done, total):
                pct = int(done / total * 100) if total else 100
                pct = min(pct, 100)
                read_bar.progress(pct / 100, text=f"Reading your document... {done}/{total} ({pct}%)")

            text = extract_text(temp_path, debug=True, progress_callback=_read_progress)
            st.session_state["last_file_hash"] = file_hash
            st.session_state["last_extracted_text"] = text
            read_bar.progress(1.0, text=f"Extracted {len(text)} characters.")

        gen_bar = st.progress(0, text="Generating questions... 0%")

        def _gen_progress(done, total):
            pct = int(done / total * 100) if total else 100
            pct = min(pct, 100)
            gen_bar.progress(pct / 100, text=f"Generating questions... {done}/{total} ({pct}%)")

        result = generate_questions_chunked(
            text, num_objective=num_objective, num_theory=num_theory,
            checkpoint_path=os.path.join(CACHE_DIR, "checkpoint.json"),
            progress_callback=_gen_progress,
        )

        with st.spinner("Building the exam document..."):
            exam_path = os.path.join(CACHE_DIR, "exam.docx")
            key_path = os.path.join(CACHE_DIR, "answer_key.docx")
            render_exam(result, header_info=header_info,
                        objective_instructions=objective_instructions,
                        theory_instructions=theory_instructions,
                        output_path=exam_path)
            render_answer_key(result, header_info=header_info, output_path=key_path)

        # Stash results in session_state so the download section below
        # (which always renders, even on reruns triggered by a download
        # button click) keeps showing them until the page is refreshed.
        st.session_state["exam_result"] = result
        st.session_state["exam_path"] = exam_path
        st.session_state["key_path"] = key_path

    except Exception as e:
        st.error("Something went wrong while generating the exam.")
        st.exception(e)
        st.stop()

# This block is OUTSIDE `if submitted`, so it renders on every rerun —
# including the rerun a download-button click triggers — as long as a
# result exists in session_state. It only clears when you refresh the
# page (new session) or click "Generate exam" again (overwrites it).
if "exam_result" in st.session_state:
    result = st.session_state["exam_result"]
    got_obj = len(result.get("section_1_objective", []))
    got_theory = len(result.get("section_2_theory", []))
    st.success(f"Generated {got_obj} objective + {got_theory} theory questions.")

    col5, col6 = st.columns(2)
    with open(st.session_state["exam_path"], "rb") as f:
        col5.download_button("Download exam (.docx)", f, file_name="exam.docx", key="dl_exam")
    with open(st.session_state["key_path"], "rb") as f:
        col6.download_button("Download answer key (.docx)", f, file_name="answer_key.docx", key="dl_key")
