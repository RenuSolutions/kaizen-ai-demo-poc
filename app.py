import io
from io import BytesIO
from typing import List, Tuple

import streamlit as st
from pptx import Presentation
from docx import Document

from openai import OpenAI
from openai import AuthenticationError, RateLimitError, APIConnectionError, APIStatusError


# -----------------------------
# Page config
# -----------------------------
st.set_page_config(
    page_title="Kaizen AI – Communication Package Generator",
    layout="wide",
)

st.title("Kaizen Deck → Communication Documents (AI Demo)")
st.caption(
    "Upload a Kaizen report-out PPTX and generate exec-ready communication artifacts. "
    "This demo generates **one document at a time** to control cost and avoid quota issues."
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
    return text[:max_chars] + "\n\n[TRUNCATED FOR DEMO COST CONTROL]"


def create_docx_bytes(title: str, body: str) -> bytes:
    """Create a .docx in memory and return bytes."""
    doc = Document()
    doc.add_heading(title, level=1)

    # Light formatting: treat lines starting with '-' or '•' as bullet-ish
    for line in body.split("\n"):
        line = line.rstrip()
        if not line.strip():
            continue
        doc.add_paragraph(line.strip())

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def generate_doc(doc_name: str, instruction: str, slide_text: str) -> str:
    """Call OpenAI and return generated text."""
    prompt = f"""
You are an expert in Lean, Kaizen, and enterprise change management.

Write: {doc_name}

Requirements:
- {instruction}
- Use clear headers and bullet points where appropriate
- Keep it concise, executive-ready, and specific to the deck
- Do NOT invent metrics; if numbers are missing, mark as "TBD" and list what is needed

Kaizen Slide Content:
{slide_text}
""".strip()

    # Use the newer Responses API (recommended)
    resp = client.responses.create(
        model="gpt-4o-mini",
        input=prompt,
    )
    return resp.output_text


# -----------------------------
# Document catalog
# -----------------------------
DOCS: List[Tuple[str, str]] = [
    ("Executive Summary", "One page max. Audience: ELT. Include: problem, root cause, solution, impact, asks/decisions."),
    ("Leader Talking Points", "One page max. Bullet-point talking points for a leader to present the Kaizen."),
    ("Change Management Summary", "Impacted roles, training needs, comms plan, adoption risks, mitigations."),
    ("Kaizen Wins & Benefits", "Benefits realization summary. Quantified wins if present; otherwise mark TBD + data needed."),
    ("30/60/90 Day Follow-Up", "30/60/90 checkpoints with owners, due dates (relative), measures, and risks."),
    ("Sustainment & Control Plan", "Controls, KPIs, cadence, auditing approach, ownership, escalation path."),
    ("Recognition Message", "A recognition/celebration message template (email/Teams post) for contributors."),
]


# -----------------------------
# UI
# -----------------------------
pptx_file = st.file_uploader("Upload Kaizen Report-Out Deck (PPTX)", type=["pptx"])

if not pptx_file:
    st.info("Upload a PPTX to begin.")
    st.stop()

pptx_bytes = pptx_file.read()
st.success("Deck uploaded successfully.")

with st.spinner("Extracting slide content..."):
    slide_text_full = extract_slide_text(pptx_bytes)

# Cost control for large decks
st.subheader("Cost Controls")
max_chars = st.slider(
    "Max characters of slide text sent to AI (controls cost)",
    min_value=5_000,
    max_value=60_000,
    value=20_000,
    step=5_000,
)
slide_text = truncate_text(slide_text_full, max_chars=max_chars)

colA, colB = st.columns([1, 1])
with colA:
    with st.expander("Preview extracted slide text (truncated view)"):
        st.text(slide_text[:8000])

with colB:
    st.warning(
        "This demo generates **ONE document at a time** to control cost and avoid quota issues.\n\n"
        "If you hit `insufficient_quota`, enable billing or increase your usage limits on the OpenAI platform."
    )

st.divider()
st.subheader("Generate One Document")

selected_doc = st.selectbox("Select ONE document to generate", [d[0] for d in DOCS])
doc_name, doc_instruction = next(d for d in DOCS if d[0] == selected_doc)

if st.button("Generate Selected Document"):
    try:
        with st.spinner(f"Generating: {doc_name} ..."):
            result_text = generate_doc(doc_name, doc_instruction, slide_text)

        st.success(f"Generated: {doc_name}")

        st.download_button(
            label=f"Download {doc_name}.docx",
            data=create_docx_bytes(doc_name, result_text),
            file_name=f"{doc_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        st.markdown("### Preview")
        st.write(result_text)

    except AuthenticationError:
        st.error(
            "OpenAI authentication failed. Confirm your `OPENAI_API_KEY` in Streamlit Secrets is correct "
            "and that you revoked any previously exposed keys."
        )
    except RateLimitError:
        st.error(
            "OpenAI quota/rate limit hit (insufficient_quota or too many requests). "
            "Check billing/usage limits on platform.openai.com, then try again."
        )
    except APIConnectionError:
        st.error("Network/API connection error. Try again; if it persists, your network may be blocking outbound calls.")
    except APIStatusError as e:
        st.error(f"OpenAI API returned an error: {e}")
    except Exception as e:
        st.error(f"Unexpected error: {e}")

