import io
import unicodedata
from io import BytesIO

import streamlit as st
from pptx import Presentation
from docx import Document
from openai import OpenAI
from openai import AuthenticationError, RateLimitError, APIConnectionError, APIStatusError


# ============================================================
# Page configuration
# ============================================================
st.set_page_config(
    page_title="Kaizen AI – Executive Summary Generator",
    layout="wide",
)

st.title("Kaizen Deck → Executive Summary (Communication-Ready)")
st.caption(
    "Upload a Kaizen report-out PPTX and generate a PPC-formatted, "
    "executive-ready one-page summary using the approved Word template."
)


# ============================================================
# Secrets / OpenAI client
# ============================================================
api_key = st.secrets.get("OPENAI_API_KEY", "").strip()
if not api_key:
    st.error(
        "Missing OPENAI_API_KEY.\n\n"
        "Add it in Streamlit:\n"
        "Manage app → Settings → Secrets\n\n"
        'OPENAI_API_KEY="sk-..."'
    )
    st.stop()

client = OpenAI(api_key=api_key)


# ============================================================
# Utilities
# ============================================================
def normalize_text(text: str) -> str:
    """
    Remove smart quotes, bullets, em-dashes, and other non-ASCII characters
    that break python-docx XML, while preserving readability.
    """
    return (
        unicodedata.normalize("NFKD", text)
        .encode("ascii", "ignore")
        .decode("ascii")
    )


def extract_slide_text(pptx_bytes: bytes) -> str:
    """Extract readable text from all PPT slides."""
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides_out = []

    for i, slide in enumerate(prs.slides, start=1):
        parts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                txt = shape.text.strip()
                if txt:
                    parts.append(txt)
        if parts:
            slides_out.append(f"Slide {i}:\n" + "\n".join(parts))

    return "\n\n".join(slides_out)


# ============================================================
# Word template helpers
# ============================================================
def replace_single(doc: Document, start_heading: str, end_heading: str, content: str):
    """
    Replace text between two headings with a paragraph block.
    """
    recording = False
    for p in doc.paragraphs:
        if start_heading in p.text:
            recording = True
            continue
        if end_heading in p.text:
            break
        if recording:
            p.text = ""

    for p in doc.paragraphs:
        if start_heading in p.text:
            idx = doc.paragraphs.index(p)
            doc.paragraphs[idx + 1].add_run(content)
            break


def replace_block(doc: Document, start_heading: str, end_heading: str, bullets: list[str]):
    """
    Replace text between headings with bullet points.
    """
    recording = False
    for p in doc.paragraphs:
        if start_heading in p.text:
            recording = True
            continue
        if end_heading in p.text:
            break
        if recording:
            p.text = ""

    for p in doc.paragraphs:
        if start_heading in p.text:
            idx = doc.paragraphs.index(p)
            insert_at = idx + 1
            for b in bullets:
                doc.paragraphs[insert_at].add_run(b)
                insert_at += 1
            break


# ============================================================
# AI Generation (STRICT PPC FORMAT)
# ============================================================
def generate_exec_summary(slide_text: str) -> dict:
    """
    Generate content EXACTLY aligned to the PPC Executive Summary template.
    """

    prompt = f"""
You are preparing a FINAL, EXECUTIVE-READY, ONE-PAGE document.

STRICT RULES:
- Follow the section structure EXACTLY as defined below
- Use concise, professional, executive language
- NO emojis, NO markdown symbols, NO headings outside the template
- Bullet points must be short and factual
- Do not invent metrics; mark missing data as TBD

SECTIONS TO RETURN (JSON ONLY):
- overview (paragraph)
- challenges (list of bullets)
- improvements (list of bullets)
- benefits (list of bullets)
- plan_0_30 (list of bullets)
- plan_30_90 (list of bullets)
- plan_6_12 (list of bullets)
- summary (paragraph)

KAIZEN SLIDE CONTENT:
{slide_text}
""".strip()

  resp = client.responses.create(
    model="gpt-4o-mini",
    input=prompt
    )
text = resp.output_text


    return resp.output_parsed


# ============================================================
# Build Word document from PPC template
# ============================================================
def build_exec_summary_doc(template_bytes: bytes, content: dict) -> bytes:
    doc = Document(BytesIO(template_bytes))

    replace_single(
        doc,
        "Overview",
        "Key Challenges Identified",
        normalize_text(content["overview"]),
    )

    replace_block(
        doc,
        "Key Challenges Identified",
        "Future-State Improvements",
        [normalize_text(x) for x in content["challenges"]],
    )

    replace_block(
        doc,
        "Future-State Improvements",
        "Organizational Benefits",
        [normalize_text(x) for x in content["improvements"]],
    )

    replace_block(
        doc,
        "Organizational Benefits",
        "Implementation Plan",
        [normalize_text(x) for x in content["benefits"]],
    )

    replace_block(
        doc,
        "Implementation Plan",
        "Summary",
        (
            [normalize_text(x) for x in content["plan_0_30"]]
            + [normalize_text(x) for x in content["plan_30_90"]]
            + [normalize_text(x) for x in content["plan_6_12"]]
        ),
    )

    replace_single(
        doc,
        "Summary",
        None,
        normalize_text(content["summary"]),
    )

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# UI
# ============================================================
pptx_file = st.file_uploader("Upload Kaizen Report-Out Deck (PPTX)", type=["pptx"])
template_file = st.file_uploader(
    "Upload PPC Executive Summary Word Template (.docx)",
    type=["docx"],
)

if not pptx_file or not template_file:
    st.info("Upload both the PPTX and the approved Word template to continue.")
    st.stop()

pptx_bytes = pptx_file.read()
template_bytes = template_file.read()

st.success("Files uploaded successfully.")

with st.spinner("Extracting slide content..."):
    slide_text = extract_slide_text(pptx_bytes)

if st.button("Generate Executive Summary"):
    try:
        with st.spinner("Generating executive summary using PPC template..."):
            content = generate_exec_summary(slide_text)
            doc_bytes = build_exec_summary_doc(template_bytes, content)

        st.success("Executive Summary generated successfully.")

        st.download_button(
            "Download Executive Summary (Word)",
            data=doc_bytes,
            file_name="Executive_Summary_PPC.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except AuthenticationError:
        st.error("OpenAI authentication failed. Check your API key.")
    except RateLimitError:
        st.error("OpenAI quota exceeded. Check billing or usage limits.")
    except APIConnectionError:
        st.error("Network error connecting to OpenAI.")
    except APIStatusError as e:
        st.error(f"OpenAI API error: {e}")
    except Exception as e:
        st.error(f"Unexpected error: {e}")
