import io
import json
import base64
from io import BytesIO
from typing import List, Tuple, Dict, Any

import streamlit as st
from pptx import Presentation
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from openai import OpenAI
from openai import AuthenticationError, RateLimitError, APIConnectionError, APIStatusError


# ============================================================
# PPC EXECUTIVE SUMMARY TEMPLATE (EMBEDDED)
# DO NOT MODIFY — this is your uploaded Word document
# ============================================================
PPC_EXEC_SUMMARY_TEMPLATE_B64 = """
UEsDBBQAAAAIAK9s0lT4xYJ8gAEAAHkFAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbK2Uy27CMBBF9/2K
hVt0K5NqEJjYpEo0qZ0i0BKKC0lF6C9sQHk3v+fZrR4Gk4Q8c+f0c8w9c5g0Fh9xq2bW8Y3x4nqz1F
... (FULL BASE64 STRING CONTINUES — KEEP AS IS) ...
UEsFBgAAAAAqACoA3gwAAH8xAAAAAA==
""".strip()


# -----------------------------
# Page config
# -----------------------------
st.set_page_config(
    page_title="Kaizen AI – Executive Summary Generator",
    layout="wide",
)

st.title("Kaizen Deck → Executive Summary (Communication-Ready)")
st.caption(
    "Upload a Kaizen report-out PPTX and generate a PPC-formatted, executive-ready "
    "one-page summary using the approved Word template."
)


# -----------------------------
# Secrets / OpenAI client
# -----------------------------
api_key = st.secrets.get("OPENAI_API_KEY", "").strip()
if not api_key:
    st.error(
        "Missing OPENAI_API_KEY.\n\n"
        "Add it in Streamlit → Settings → Secrets:\n"
        'OPENAI_API_KEY="sk-..."'
    )
    st.stop()

client = OpenAI(api_key=api_key)


# -----------------------------
# PPTX text extraction
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
    return text if len(text) <= max_chars else text[:max_chars]


# -----------------------------
# Word helpers (template injection)
# -----------------------------
def load_template() -> Document:
    raw = base64.b64decode(PPC_EXEC_SUMMARY_TEMPLATE_B64)
    return Document(BytesIO(raw))


def iter_paragraphs(doc: Document):
    return list(doc.paragraphs)


def delete_paragraph(p: Paragraph):
    p._element.getparent().remove(p._element)


def insert_after(p: Paragraph, text: str, style):
    new_p = OxmlElement("w:p")
    p._element.addnext(new_p)
    para = Paragraph(new_p, p._parent)
    para.style = style
    para.add_run(text)
    return para


def find_heading_index(paragraphs, heading):
    for i, p in enumerate(paragraphs):
        if p.text.strip() == heading:
            return i
    return -1


def replace_block(doc, start, end, lines):
    paras = iter_paragraphs(doc)
    s = find_heading_index(paras, start)
    e = find_heading_index(paras, end)

    if s == -1 or e == -1 or e <= s:
        return

    block = paras[s+1:e]
    style = block[0].style if block else paras[s].style

    for p in block:
        delete_paragraph(p)

    paras = iter_paragraphs(doc)
    anchor = paras[s]

    last = anchor
    for line in lines:
        last = insert_after(last, line, style)


def replace_single(doc, start, end, text):
    paras = iter_paragraphs(doc)
    s = find_heading_index(paras, start)
    e = find_heading_index(paras, end)

    if s == -1:
        return

    block = paras[s+1:e] if e != -1 else paras[s+1:]
    style = block[0].style if block else paras[s].style

    for p in block:
        delete_paragraph(p)

    paras = iter_paragraphs(doc)
    insert_after(paras[s], text, style)


def set_title(doc, title):
    for p in doc.paragraphs:
        if p.text.startswith("Executive Summary"):
            p.text = title
            return


def build_exec_summary(deck_name: str, content: Dict[str, Any]) -> bytes:
    doc = load_template()

    set_title(doc, f"Executive Summary – {deck_name} (One Page)")

    replace_single(doc, "Overview", "Key Challenges Identified", content["overview"])

    replace_block(
        doc,
        "Key Challenges Identified",
        "Future-State Improvements",
        content["challenges"],
    )

    replace_block(
        doc,
        "Future-State Improvements",
        "Organizational Benefits",
        content["improvements"],
    )

    replace_block(
        doc,
        "Organizational Benefits",
        "Implementation Plan",
        content["benefits"],
    )

    replace_block(
        doc,
        "Implementation Plan",
        "Summary",
        [
            f"0–30 Days: {content['plan_0_30']}",
            f"30–90 Days: {content['plan_30_90']}",
            f"6–12 Months: {content['plan_6_12']}",
        ],
    )

    replace_single(doc, "Summary", "", content["summary"])

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# -----------------------------
# AI generation (structured)
# -----------------------------
def generate_structured_summary(slide_text: str) -> Dict[str, Any]:
    prompt = f"""
Return ONLY valid JSON with these keys:
overview, challenges, improvements, benefits,
plan_0_30, plan_30_90, plan_6_12, summary

Rules:
- challenges / improvements / benefits are arrays of short bullets (no hyphens)
- Do NOT invent metrics; use "TBD" if missing
- No extra text outside JSON

KAIZEN CONTENT:
{slide_text}
""".strip()

    r = client.responses.create(
        model="gpt-4o-mini",
        input=prompt,
    )

    txt = r.output_text.strip()
    return json.loads(txt[txt.find("{"): txt.rfind("}")+1])


# -----------------------------
# UI
# -----------------------------
pptx_file = st.file_uploader("Upload Kaizen Report-Out Deck (PPTX)", type=["pptx"])
if not pptx_file:
    st.stop()

pptx_bytes = pptx_file.read()
slide_text = truncate_text(extract_slide_text(pptx_bytes), 20000)

deck_name = pptx_file.name.replace(".pptx", "")

if st.button("Generate Executive Summary"):
    try:
        with st.spinner("Generating executive-ready document…"):
            structured = generate_structured_summary(slide_text)
            docx = build_exec_summary(deck_name, structured)

        st.success("Executive Summary ready")

        st.download_button(
            "Download Executive Summary.docx",
            data=docx,
            file_name="Executive Summary.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        st.error(str(e))


