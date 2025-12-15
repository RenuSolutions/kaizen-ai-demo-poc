import io
import json
import re
from io import BytesIO
from typing import Dict, List, Optional

import streamlit as st
from pptx import Presentation
from docx import Document
from docx.text.paragraph import Paragraph

from openai import OpenAI
from openai import AuthenticationError, RateLimitError, APIConnectionError, APIStatusError


# -----------------------------
# Page config
# -----------------------------
st.set_page_config(
    page_title="Kaizen Executive Summary (Communication-Ready)",
    layout="wide",
)

st.title("Kaizen Deck – Executive Summary (Communication-Ready)")
st.caption(
    "Upload a Kaizen report-out PPTX and generate a PPC-formatted executive-ready one-page summary "
    "using the approved Word template placeholders."
)

# -----------------------------
# Secrets / client setup
# -----------------------------
api_key = st.secrets.get("OPENAI_API_KEY", "").strip()
if not api_key:
    st.warning(
        "Missing `OPENAI_API_KEY`.\n\n"
        "Add it in Manage app → Settings → Secrets as:\n\n"
        'OPENAI_API_KEY="sk-..."'
    )
    st.stop()

client = OpenAI(api_key=api_key)

# -----------------------------
# Helpers
# -----------------------------
def extract_slide_text(pptx_bytes: bytes) -> str:
    """Extract plain text from all slides."""
    prs = Presentation(io.BytesIO(pptx_bytes))
    out = []
    for i, slide in enumerate(prs.slides, start=1):
        parts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                t = shape.text.strip()
                if t:
                    parts.append(t)
        if parts:
            out.append(f"Slide {i}:\n" + "\n".join(parts))
    return "\n\n".join(out)


def truncate_text(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + "\n\n[TRUNCATED FOR COST CONTROL]"


def _clean_bullets(items: List[str]) -> List[str]:
    cleaned = []
    for x in items or []:
        if not x:
            continue
        s = str(x).strip()
        s = re.sub(r"^[•\-\*\u2022]\s*", "", s)  # strip leading bullet chars
        if s:
            cleaned.append(s)
    return cleaned


def _safe_get(d: Dict, key: str, default):
    v = d.get(key, default)
    return v if v is not None else default


def generate_structured_sections(slide_text: str) -> Dict:
    """
    Ask the model for JSON only, then parse.
    Output keys:
      - overview: string
      - challenges: [string]
      - improvements: [string]
      - benefits: [string]
      - plan: [string]  (bullets including 0–30, 30–90, 6–12 mo as applicable)
      - summary: string
    """
    prompt = f"""
You are producing content for a one-page Executive Summary formatted for leadership.

Return ONLY valid JSON (no markdown). Use this exact schema:
{{
  "overview": "1 short paragraph",
  "challenges": ["bullet", "bullet", "bullet"],
  "improvements": ["bullet", "bullet", "bullet"],
  "benefits": ["bullet", "bullet", "bullet"],
  "plan": ["0–30 Days: ...", "30–90 Days: ...", "6–12 Months: ..."],
  "summary": "1 short paragraph"
}}

Rules:
- Keep it concise and communication-ready.
- Do not invent numbers; if missing, write "TBD" and specify what data is needed.
- Base everything strictly on the slide content.

SLIDE CONTENT:
{slide_text}
""".strip()

    resp = client.responses.create(
        model="gpt-4o-mini",
        input=prompt,
    )
    text = (resp.output_text or "").strip()

    # Try strict JSON parse; if model returns extra text, extract first JSON object.
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r"\{.*\}", text, flags=re.S)
        if not m:
            raise ValueError("Model did not return JSON.")
        return json.loads(m.group(0))


def iter_all_paragraphs(doc: Document):
    """Yield all paragraphs in document, including those in tables."""
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def _replace_paragraph_text_preserve_style(paragraph: Paragraph, new_text: str):
    """
    Replace text while preserving paragraph style.
    If paragraph has runs, keep formatting of the first run.
    """
    new_text = new_text or ""
    if paragraph.runs:
        # Put all text into first run, clear the rest
        paragraph.runs[0].text = new_text
        for r in paragraph.runs[1:]:
            r.text = ""
    else:
        paragraph.text = new_text


def _insert_paragraph_after(paragraph: Paragraph, text: str) -> Paragraph:
    """Insert a new paragraph after `paragraph` with given text."""
    new_p = paragraph.insert_paragraph_after(text)
    return new_p


def fill_placeholders_in_template(template_bytes: bytes, content: Dict) -> bytes:
    """
    Replace placeholders in the Word template:
      {{OVERVIEW}}, {{CHALLENGES}}, {{IMPROVEMENTS}}, {{BENEFITS}}, {{PLAN}}, {{SUMMARY}}
    Bullet sections are inserted as bullet paragraphs using the placeholder paragraph style.
    """
    doc = Document(BytesIO(template_bytes))

    # Normalize content
    overview = str(_safe_get(content, "overview", "")).strip()
    summary = str(_safe_get(content, "summary", "")).strip()

    challenges = _clean_bullets(_safe_get(content, "challenges", []))
    improvements = _clean_bullets(_safe_get(content, "improvements", []))
    benefits = _clean_bullets(_safe_get(content, "benefits", []))
    plan = _clean_bullets(_safe_get(content, "plan", []))

    placeholder_map = {
        "{{OVERVIEW}}": ("para", overview),
        "{{SUMMARY}}": ("para", summary),
        "{{CHALLENGES}}": ("bullets", challenges),
        "{{IMPROVEMENTS}}": ("bullets", improvements),
        "{{BENEFITS}}": ("bullets", benefits),
        "{{PLAN}}": ("bullets", plan),
    }

    # Replace each placeholder occurrence
    for p in iter_all_paragraphs(doc):
        raw = (p.text or "").strip()
        if raw in placeholder_map:
            kind, value = placeholder_map[raw]

            if kind == "para":
                _replace_paragraph_text_preserve_style(p, value)

            elif kind == "bullets":
                items: List[str] = value if isinstance(value, list) else []
                if not items:
                    _replace_paragraph_text_preserve_style(p, "TBD")
                else:
                    # Use the placeholder paragraph's style for bullets
                    bullet_style = p.style

                    # Replace placeholder paragraph with first bullet
                    _replace_paragraph_text_preserve_style(p, items[0])
                    p.style = bullet_style

                    # Insert remaining bullets as new paragraphs after
                    anchor = p
                    for item in items[1:]:
                        new_p = _insert_paragraph_after(anchor, item)
                        new_p.style = bullet_style
                        anchor = new_p

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


# -----------------------------
# UI
# -----------------------------
left, right = st.columns([1, 1])

with left:
    pptx_file = st.file_uploader("Upload Kaizen Report-Out Deck (PPTX)", type=["pptx"])
with right:
    template_file = st.file_uploader("Upload PPC Executive Summary Word Template (.docx)", type=["docx"])

st.info(
    "Template must include placeholders on their own lines: "
    "{{OVERVIEW}}, {{CHALLENGES}}, {{IMPROVEMENTS}}, {{BENEFITS}}, {{PLAN}}, {{SUMMARY}}"
)

if not pptx_file or not template_file:
    st.stop()

pptx_bytes = pptx_file.read()
template_bytes = template_file.read()
st.success("Files uploaded successfully.")

st.subheader("Cost Controls")
max_chars = st.slider(
    "Max characters of slide text sent to the model",
    min_value=5_000,
    max_value=60_000,
    value=20_000,
    step=5_000,
)

with st.spinner("Extracting slide content..."):
    slide_text_full = extract_slide_text(pptx_bytes)
slide_text = truncate_text(slide_text_full, max_chars=max_chars)

with st.expander("Preview extracted slide text (truncated)"):
    st.text(slide_text[:8000])

st.divider()

if st.button("Generate Executive Summary"):
    try:
        with st.spinner("Generating structured content..."):
            sections = generate_structured_sections(slide_text)

        with st.spinner("Filling Word template..."):
            final_docx = fill_placeholders_in_template(template_bytes, sections)

        st.success("Executive Summary generated.")

        st.download_button(
            label="Download Executive Summary (PPC Template).docx",
            data=final_docx,
            file_name="Executive Summary - PPC Template.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        st.subheader("Preview (Structured Content)")
        st.json(sections)

    except AuthenticationError:
        st.error("OpenAI authentication failed. Check the OPENAI_API_KEY in Streamlit Secrets.")
    except RateLimitError:
        st.error("Quota/rate limit hit. Check billing/usage limits in your OpenAI account.")
    except APIConnectionError:
        st.error("Network/API connection error. Try again.")
    except APIStatusError as e:
        st.error(f"OpenAI API returned an error: {e}")
    except Exception as e:
        st.error(f"Unexpected error: {e}")

