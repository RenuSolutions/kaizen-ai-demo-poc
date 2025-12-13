import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import os

# ---------- Page Config ----------
st.set_page_config(
    page_title="Kaizen AI – Communication Package Generator",
    layout="wide"
)

st.title("Kaizen Deck → Communication Documents (AI Demo)")
st.caption(
    "Proof of Concept: Upload a Kaizen report-out deck and auto-generate "
    "executive-ready communication artifacts."
)

# ---------- Secrets Check ----------
missing = []
if "OPENAI_API_KEY" not in st.secrets:
    missing.append("OPENAI_API_KEY")

if missing:
    st.warning(
        "⚠️ Missing secrets: " + ", ".join(missing) +
        ". Add them in **Manage app → Settings → Secrets**."
    )
    st.stop()

# ---------- OpenAI Setup ----------
from openai import OpenAI
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# ---------- Helper Functions ----------
def extract_slide_text(pptx_file) -> str:
    prs = Presentation(pptx_file)
    slides_text = []

    for i, slide in enumerate(prs.slides, start=1):
        slide_content = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                slide_content.append(shape.text)
        if slide_content:
            slides_text.append(f"Slide {i}:\n" + "\n".join(slide_content))

    return "\n\n".join(slides_text)


def generate_doc(title, instruction, slide_text):
    prompt = f"""
You are an expert in Lean, Kaizen, and enterprise change management.

Using the Kaizen slide content below, create:
{instruction}

Kaizen Slides:
{slide_text}
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
    )
    return response.choices[0].message.content


# ---------- UI ----------
uploaded_file = st.file_uploader(
    "Upload Kaizen Report-Out Deck (PPTX)",
    type=["pptx"]
)

if uploaded_file:
    st.success("Deck uploaded successfully")

    with st.spinner("Extracting slide content..."):
        slide_text = extract_slide_text(uploaded_file)

    st.subheader("Preview: Extracted Slide Content")
    with st.expander("Show extracted text"):
        st.text(slide_text[:5000])

    if st.button("Generate Communication Package"):
        with st.spinner("Generating documents with AI..."):
            outputs = {}

            outputs["Executive Summary"] = generate_doc(
                "Executive Summary",
                "A one-page executive summary suitable for ELT review.",
                slide_text
            )

            outputs["Leader Talking Points"] = generate_doc(
                "Leader Talking Points",
                "A concise one-page set of talking points for leaders.",
                slide_text
            )

            outputs["Change Management Summary"] = generate_doc(
                "Change Management Summary",
                "A clear change management summary including impacts, risks, and mitigation.",
                slide_text
            )

            outputs["Wins & Benefits"] = generate_doc(
                "Kaizen Wins & Benefits",
                "A benefits realization summary highlighting measurable wins.",
                slide_text
            )

            outputs["30-60-90 Day Follow-Up"] = generate_doc(
                "30/60/90 Day Follow-Up",
                "A structured 30/60/90 day follow-up plan.",
                slide_text
            )

            outputs["Sustainment Plan"] = generate_doc(
                "Sustainment & Control Plan",
                "A sustainment plan including controls, owners, and cadence.",
                slide_text
            )

        st.success("Documents generated successfully")

        # ---------- Display Outputs ----------
        for name, content in outputs.items():
            st.subheader(name)
            st.text_area(
                label=name,
                value=content,
                height=300
            )

        st.info(
            "Next phase: Persist outputs to Snowflake stages & tables "
            "and expose via Snowflake Streamlit portal."
        )

else:
    st.info("Upload a Kaizen PPTX to begin.")

