import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from datetime import date

import streamlit as st
from dotenv import load_dotenv
from google import genai
from groq import Groq
from docx import Document
from docx.shared import Pt
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

load_dotenv()

st.set_page_config(page_title="LegalEase AI", page_icon="⚖️", layout="wide")
st.title("⚖️ LegalEase AI")
st.caption("AI-assisted legal document drafting with Gemini and Groq.")

st.info(
    "This application creates drafts for educational/general-information use. "
    "It is not legal advice. Have an appropriately qualified lawyer review documents "
    "before real-world use."
)

GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")
GROQ_MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-120b")


def get_api_keys():
    return {
        "gemini": os.getenv("GEMINI_API_KEY", "").strip(),
        "groq": os.getenv("GROQ_API_KEY", "").strip(),
    }


def build_prompt(doc_type, parties, terms, effective_date):
    return f"""
You are a legal-document drafting assistant.
Create a clear, structured FIRST DRAFT of a {doc_type}.

Inputs:
Parties: {parties}
Effective date: {effective_date}
Terms and conditions: {terms}

Requirements:
- Use headings and numbered clauses.
- Do not invent names, addresses, prices, dates or obligations that were not provided.
- If important information is missing, mark it as [TO BE COMPLETED].
- Use neutral professional language.
- Include signature sections where appropriate.
- Add a short "Drafting Notice" at the end saying the document should be reviewed for applicable local law.
- This is drafting assistance, not legal advice.
Return only the document text.
""".strip()


def generate_with_gemini(prompt):
    key = get_api_keys()["gemini"]
    if not key or key == "PASTE_YOUR_GEMINI_API_KEY_HERE":
        raise ValueError("GEMINI_API_KEY is missing in .env")

    client = genai.Client(api_key=key)
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
    )
    return response.text


def generate_with_groq(prompt):
    key = get_api_keys()["groq"]
    if not key or key == "PASTE_YOUR_GROQ_API_KEY_HERE":
        raise ValueError("GROQ_API_KEY is missing in .env")

    client = Groq(api_key=key)
    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a legal-document drafting assistant. "
                    "Return only the requested document text."
                ),
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0.2,
        max_completion_tokens=12000,
    )
    return response.choices[0].message.content


def generate_with_both(prompt):
    """Call Gemini and Groq concurrently with the same prompt."""
    results = {}

    with ThreadPoolExecutor(max_workers=2) as executor:
        futures = {
            executor.submit(generate_with_gemini, prompt): "Gemini",
            executor.submit(generate_with_groq, prompt): "Groq",
        }

        for future in as_completed(futures):
            provider = futures[future]
            try:
                results[provider] = {"ok": True, "text": future.result()}
            except Exception as exc:
                results[provider] = {"ok": False, "error": str(exc)}

    return results


def make_docx(text, title):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)

    p = doc.add_paragraph()
    r = p.add_run(title.upper())
    r.bold = True
    r.font.size = Pt(16)

    for line in text.splitlines():
        doc.add_paragraph(line)

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def make_pdf(text, title):
    out = BytesIO()
    pdf = SimpleDocTemplate(
        out,
        pagesize=A4,
        rightMargin=50,
        leftMargin=50,
        topMargin=50,
        bottomMargin=50,
    )
    styles = getSampleStyleSheet()
    story = [Paragraph(title.upper(), styles["Title"]), Spacer(1, 12)]

    for line in text.splitlines():
        safe = (
            line.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        if not safe.strip():
            story.append(Spacer(1, 8))
        elif safe.isupper() and len(safe) < 70:
            story.append(Paragraph(f"<b>{safe}</b>", styles["Heading2"]))
        else:
            story.append(Paragraph(safe, styles["BodyText"]))

    pdf.build(story)
    return out.getvalue()


with st.sidebar:
    st.header("Document Details")

    doc_type = st.selectbox(
        "Document type",
        [
            "Agreement",
            "NDA (Non-Disclosure Agreement)",
            "Lease Agreement",
            "Employment Contract",
            "Employment Offer Letter",
            "Freelance Contract",
        ],
    )

    parties = st.text_area(
        "Parties involved",
        placeholder="Example: Jane Doe (Service Provider), ABC Ltd (Client)",
    )

    terms = st.text_area(
        "Terms & conditions",
        placeholder=(
            "Example: Payment within 30 days; Confidentiality; "
            "15 days notice for termination"
        ),
    )

    effective_date = st.date_input("Effective date", value=date.today())

    provider = st.radio(
        "AI provider",
        ["Gemini", "Groq", "Both (run simultaneously)"],
    )

    generate = st.button(
        "✨ Generate",
        type="primary",
        use_container_width=True,
    )

keys = get_api_keys()
with st.sidebar:
    st.divider()
    st.caption(f"Gemini key: {'✅ configured' if keys['gemini'] else '❌ missing'}")
    st.caption(f"Groq key: {'✅ configured' if keys['groq'] else '❌ missing'}")

if "document" not in st.session_state:
    st.session_state.document = ""

if "gemini_result" not in st.session_state:
    st.session_state.gemini_result = ""

if "groq_result" not in st.session_state:
    st.session_state.groq_result = ""

if generate:
    if not parties.strip() or not terms.strip():
        st.error("Please enter the parties and terms.")
    else:
        prompt = build_prompt(
            doc_type,
            parties,
            terms,
            effective_date.strftime("%d/%m/%Y"),
        )

        if provider == "Gemini":
            with st.spinner("Generating with Gemini..."):
                try:
                    st.session_state.document = generate_with_gemini(prompt)
                    st.session_state.gemini_result = st.session_state.document
                    st.success("Draft generated with Gemini.")
                except Exception as exc:
                    st.error(f"Gemini generation failed: {exc}")

        elif provider == "Groq":
            with st.spinner("Generating with Groq..."):
                try:
                    st.session_state.document = generate_with_groq(prompt)
                    st.session_state.groq_result = st.session_state.document
                    st.success("Draft generated with Groq.")
                except Exception as exc:
                    st.error(f"Groq generation failed: {exc}")

        else:
            with st.spinner("Gemini + Groq are generating simultaneously..."):
                results = generate_with_both(prompt)

            if results.get("Gemini", {}).get("ok"):
                st.session_state.gemini_result = results["Gemini"]["text"]
            else:
                st.session_state.gemini_result = ""
                st.error(
                    "Gemini failed: "
                    + results.get("Gemini", {}).get("error", "Unknown error")
                )

            if results.get("Groq", {}).get("ok"):
                st.session_state.groq_result = results["Groq"]["text"]
            else:
                st.session_state.groq_result = ""
                st.error(
                    "Groq failed: "
                    + results.get("Groq", {}).get("error", "Unknown error")
                )

            # Put the first successful result into the editable document.
            st.session_state.document = (
                st.session_state.gemini_result
                or st.session_state.groq_result
            )

            if st.session_state.gemini_result or st.session_state.groq_result:
                st.success("Gemini and Groq finished. Compare both outputs below.")

if provider == "Both (run simultaneously)" and (
    st.session_state.gemini_result or st.session_state.groq_result
):
    st.subheader("AI Comparison")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 🟦 Gemini")
        if st.session_state.gemini_result:
            st.text_area(
                "Gemini output",
                value=st.session_state.gemini_result,
                height=420,
                key="gemini_output_view",
            )
        else:
            st.warning("No Gemini output.")

    with col2:
        st.markdown("### 🟩 Groq")
        if st.session_state.groq_result:
            st.text_area(
                "Groq output",
                value=st.session_state.groq_result,
                height=420,
                key="groq_output_view",
            )
        else:
            st.warning("No Groq output.")

st.subheader("Editable Document")
st.session_state.document = st.text_area(
    "Edit the generated document",
    value=st.session_state.document,
    height=560,
    placeholder="Generate a document, then edit the text here.",
    key="editable_document",
)

if st.session_state.document.strip():
    text = st.session_state.document

    st.subheader("Download")
    c1, c2, c3 = st.columns(3)

    with c1:
        st.download_button(
            "⬇️ Download TXT",
            text.encode("utf-8"),
            "legalease_document.txt",
            "text/plain",
            use_container_width=True,
        )

    with c2:
        st.download_button(
            "⬇️ Download DOCX",
            make_docx(text, doc_type),
            "legalease_document.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    with c3:
        st.download_button(
            "⬇️ Download PDF",
            make_pdf(text, doc_type),
            "legalease_document.pdf",
            "application/pdf",
            use_container_width=True,
        )

st.divider()
st.caption(
    f"LegalEase AI • Gemini ({GEMINI_MODEL}) + Groq ({GROQ_MODEL})"
)
