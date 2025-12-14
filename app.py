import io
from io import BytesIO
import streamlit as st
from pptx import Presentation
from docx import Document
from openai import OpenAI
from openai import AuthenticationError, RateLimitError


# ============================
# Page Config
# ============================
st.set_page_config(
    page_title="Kaizen → Executive Summary (Communication-Ready)",
    layout="wide",
)

st.title("Kaizen Deck → Executive Summary (Communication-Ready)")
st.caption(
    "Upload a Kaizen PPTX and the approved PPC Executive Summary Word template. "
    "This app generates content ONLY and injects it into the template with branding preserved."
)


# ============================
# OpenAI Setup
# ============================
api_key = st.secrets.get("OPENAI_API_KEY", "").strip()
if not api_key:
    st.error("Missing OPENAI_API_KEY in Streamlit Secrets.")
    st.stop()

client = OpenAI(api_key=api_key)


# ============================
# Helpers
# ============================
def extract_slide_text(pptx_bytes: bytes) -> str:
    prs = Presentation(BytesIO(pptx_bytes))
    text_blocks = []

    for i, slide in enumerate(prs.slides, start=1):
        slide_text = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                slide_text.append(shape.text.strip())
        if slide_text:
            text_blocks.append(f"Slide {i}:\n" + "\n".join(slide_text))

    return "\n\n".join(text_blocks)


def generate_exec_summary(slide_text: str) -> dict:
    """
    Returns structured sections that match the PPC template exactly
    """
    prompt = f"""
You are writing a PPC Partners Executive Summary (ONE PAGE).

STRICT RULES:
- Match these sections EXACTLY
- No markdown
- No emojis
- Professional consulting tone
- Concise, leadership-ready

SECTIONS REQUIRED:
1. Overview
2. Key Challenges Identified
3. Future-State Improvements
4. Organizational Benefits
5. Implementation Plan
6. Summary

Kaizen Deck Content:
{slide_text}
"""

    resp = client.responses.create(
        model="gpt-4o-mini",
        input=prompt
    )

    text = resp.output_text.strip()

    sections = {
        "Overview": "",
        "Key Challenges Identified": "",
        "Future-State Improvements": "",
        "Organizational Benefits": "",
        "Implementation Plan": "",
        "Summary": ""
    }

    current = None
    for line in text.splitlines():
        line = line.strip()
        if line in sections:
            current = line
            continue
        if current and line:
            sections[current] += line + "\n"

    return sections


def fill_word_template(template_bytes: bytes, sections: dict) -> bytes:
    doc = Document(BytesIO(template_bytes))

    for para in doc.paragraphs:
        title = para.text.strip()
        if title in sections:
            para.clear()
            para.add_run(title)

            p = para._element.getnext()
            if p is not None:
                p.getparent().remove(p)

            for line in sections[title].splitlines():
                if line.strip():
                    doc.add_paragraph(line.strip(), style=para.style)

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================
# UI
# ============================
pptx_file = st.file_uploader("Upload Kaizen Report-Out Deck (PPTX)", type=["pptx"])
template_file = st.file_uploader(
    "Upload PPC Executive Summary Word Template (.docx)",
    type=["docx"]
)

if not pptx_file or not template_file:
    st.info("Upload BOTH the PPTX and the approved Word template to continue.")
    st.stop()

pptx_bytes = pptx_file.read()
template_bytes = template_file.read()

st.success("Files uploaded successfully.")

if st.button("Generate Executive Summary"):
    try:
        with st.spinner("Extracting slide content..."):
            slide_text = extract_slide_text(pptx_bytes)

        with st.spinner("Generating executive-ready content with AI..."):
            sections = generate_exec_summary(slide_text)

        with st.spinner("Applying content to approved PPC template..."):
            final_doc = fill_word_template(template_bytes, sections)

        st.success("Executive Summary generated successfully.")

        st.download_button(
            "Download Executive Summary (.docx)",
            data=final_doc,
            file_name="PPC_Executive_Summary.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except AuthenticationError:
        st.error("Invalid OpenAI API key.")
    except RateLimitError:
        st.error("OpenAI quota exceeded. Check billing.")
    except Exception as e:
        st.error(f"Unexpected error: {e}")

