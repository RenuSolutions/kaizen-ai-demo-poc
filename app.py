import io
import json
import re
from io import BytesIO
from typing import Dict, List, Any

import streamlit as st
from pptx import Presentation

from docx import Document
from docx.oxml import OxmlElement

from openai import OpenAI
from openai import AuthenticationError, RateLimitError, APIConnectionError, APIStatusError


# -----------------------------
# Page config
# -----------------------------
st.set_page_config(
    page_title="Kaizen Deck – Executive Summary (Communication-Ready)",
    layout="wide",
)

st.title("Kaizen Deck – Executive Summary (Communication-Ready)")
st.caption(
    "Upload a Kaizen report-out PPTX and generate a PPC-formatted, executive-ready one-page summary "
    "using the approved Word template placeholders."
)

# -----------------------------
# Secrets / client setup
# -----------------------------
api_key = st.secrets.get("OPENAI_API_KEY", "").strip()
if not api_key:
    st.warning(
        "⚠️ Missing `OPENAI_API_KEY`.\n\n"
        "Add it in **Manage app → Settings → Secrets** as:\n\n"
        '`OPENAI_API_KEY="sk-..."`'
    )
    st.stop()

client = OpenAI(api_key=api_key)


# -----------------------------
# Helpers
# -----------------------------
def extract_slide_text(pptx_bytes: bytes) -> str:
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
    return text[:max_chars] + "\n\n[TRUNCATED FOR DEMO COST CONTROL]"


def _insert_paragraph_after(paragraph, text: str = "", style: str = None):
    """
    Insert a new paragraph AFTER 'paragraph' (python-docx doesn't provide insert_paragraph_after).
    """
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)

    new_para = paragraph._parent.add_paragraph()
    new_para._p = new_p

    if style:
        try:
            new_para.style = style
        except Exception:
            pass

    if text:
        new_para.add_run(text)

    return new_para


def _clean_bullets(lines: List[str]) -> List[str]:
    cleaned = []
    for ln in lines:
        ln = ln.strip()
        ln = re.sub(r"^[-•\u2022]\s*", "", ln)
        if ln:
            cleaned.append(ln)
    return cleaned


def _normalize_sections(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Ensure keys exist and types are correct.
    """
    def as_str(x):
        return (x or "").strip()

    def as_list(x):
        if x is None:
            return []
        if isinstance(x, list):
            return [str(i).strip() for i in x if str(i).strip()]
        # if model gave a single string with lines
        if isinstance(x, str):
            return [s.strip() for s in x.splitlines() if s.strip()]
        return [str(x).strip()] if str(x).strip() else []

    return {
        "overview": as_str(data.get("overview")),
        "challenges": as_list(data.get("challenges")),
        "improvements": as_list(data.get("improvements")),
        "benefits": as_list(data.get("benefits")),
        "plan": as_list(data.get("plan")),
        "summary": as_str(data.get("summary")),
        "title": as_str(data.get("title")),  # optional
    }


def generate_exec_summary_json(slide_text: str) -> (Dict[str, Any], str):
    """
    Generate content as STRICT JSON. Returns (sections_dict, raw_output_text).
    """
    prompt = f"""
You are writing a PPC Partners executive one-page Kaizen summary.

Return ONLY valid JSON (no markdown, no backticks). The JSON MUST have exactly these keys:
- "title" (string)  -> example: "Executive Summary – <Kaizen Name> (One Page)"
- "overview" (string) -> 2–3 sentences. No bullets.
- "challenges" (array of strings) -> 3–6 bullets max
- "improvements" (array of strings) -> 3–6 bullets max
- "benefits" (array of strings) -> 3–6 bullets max
- "plan" (array of strings) -> 2–3 bullets max (use time horizons like 0–30, 30–90, 6–12 months if possible)
- "summary" (string) -> 1–2 sentences. No bullets.

Rules:
- Use only facts supported by the slide content; if missing, use "TBD" briefly.
- Keep it concise and executive-ready.
- Do not include any other keys or commentary.

Kaizen slide content:
{slide_text}
""".strip()

    resp = client.responses.create(
        model="gpt-4o-mini",
        input=prompt,
    )
    raw = resp.output_text.strip()

    # Try strict JSON parse
    try:
        data = json.loads(raw)
        return _normalize_sections(data), raw
    except Exception:
        # common recovery: sometimes the model adds leading text; try to extract the first {...}
        m = re.search(r"\{.*\}", raw, flags=re.DOTALL)
        if m:
            data = json.loads(m.group(0))
            return _normalize_sections(data), raw
        raise ValueError("AI did not return valid JSON.")


def fill_docx_template(template_bytes: bytes, sections: Dict[str, Any]) -> bytes:
    """
    Replace placeholders in the template with content.
    Placeholders must be on their own line:
    {{OVERVIEW}}, {{CHALLENGES}}, {{IMPROVEMENTS}}, {{BENEFITS}}, {{PLAN}}, {{SUMMARY}}
    """
    doc = Document(BytesIO(template_bytes))

    placeholders = {
        "{{OVERVIEW}}": "overview",
        "{{CHALLENGES}}": "challenges",
        "{{IMPROVEMENTS}}": "improvements",
        "{{BENEFITS}}": "benefits",
        "{{PLAN}}": "plan",
        "{{SUMMARY}}": "summary",
    }

    # Validate placeholders exist
    all_text = "\n".join([p.text.strip() for p in doc.paragraphs])
    missing = [ph for ph in placeholders.keys() if ph not in all_text]
    if missing:
        raise ValueError(
            "Template must include placeholders on their own lines: "
            "{{OVERVIEW}}, {{CHALLENGES}}, {{IMPROVEMENTS}}, {{BENEFITS}}, {{PLAN}}, {{SUMMARY}}"
        )

    for p in list(doc.paragraphs):
        key = p.text.strip()
        if key in placeholders:
            field = placeholders[key]
            p.text = ""  # remove placeholder

            if field in ("overview", "summary"):
                text = sections.get(field, "").strip() or "TBD"
                p.add_run(text)

            else:
                items = sections.get(field, [])
                if not items:
                    items = ["TBD"]

                items = _clean_bullets(items)

                # first bullet goes in same paragraph
                try:
                    p.style = "List Bullet"
                except Exception:
                    pass
                p.add_run(items[0])

                prev = p
                for item in items[1:]:
                    newp = _insert_paragraph_after(prev, "", style="List Bullet")
                    newp.add_run(item)
                    prev = newp

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# -----------------------------
# UI
# -----------------------------
col1, col2 = st.columns(2)

with col1:
    pptx_file = st.file_uploader("Upload Kaizen Report-Out Deck (PPTX)", type=["pptx"])

with col2:
    template_file = st.file_uploader("Upload PPC Executive Summary Word Template (.docx)", type=["docx"])

st.info("Upload both the PPTX and the approved Word template to continue.")

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

slide_text_full = extract_slide_text(pptx_bytes)
slide_text = truncate_text(slide_text_full, max_chars=max_chars)

with st.expander("Preview extracted slide text (truncated)"):
    st.text(slide_text[:12000])

st.divider()

if st.button("Generate Executive Summary"):
    try:
        with st.spinner("Generating summary content..."):
            sections, raw_ai = generate_exec_summary_json(slide_text)

        with st.expander("Show raw AI output (debug)"):
            st.text(raw_ai)

        with st.spinner("Filling Word template..."):
            output_docx = fill_docx_template(template_bytes, sections)

        st.success("Executive Summary generated.")

        st.download_button(
            label="Download Executive Summary (.docx)",
            data=output_docx,
            file_name="Executive_Summary_One_Page.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        st.subheader("Preview (what was inserted)")
        st.write(sections)

    except ValueError as ve:
        st.error(str(ve))
    except AuthenticationError:
        st.error("OpenAI authentication failed. Check your OPENAI_API_KEY in Streamlit Secrets.")
    except RateLimitError:
        st.error("OpenAI quota/rate limit hit. Check billing/usage limits, then try again.")
    except APIConnectionError:
        st.error("Network/API connection error. Try again.")
    except APIStatusError as e:
        st.error(f"OpenAI API returned an error: {e}")
    except Exception as e:
        st.error(f"Unexpected error: {e}")



