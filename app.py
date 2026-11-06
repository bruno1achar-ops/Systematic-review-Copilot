
import io
from collections import Counter

import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(
    page_title="Systematic Review Copilot",
    page_icon="📚",
    layout="wide",
)

st.title("📚 Systematic Review Copilot")
st.caption(
    "A free, template-based workspace for organizing "
    "verified evidence into manuscript drafts."
)

st.warning(
    "This app does not use an AI API. It organizes your "
    "notes and extraction data. Verify all scientific "
    "claims, calculations, and references before submission."
)

uploaded_file = st.file_uploader(
    "Upload your study extraction sheet",
    type=["xlsx", "csv"],
)

df = None

if uploaded_file:
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        df.columns = df.columns.astype(str)
        df = df.loc[:, ~df.columns.duplicated()]

        st.success(f"Loaded {len(df)} rows and {len(df.columns)} columns.")
        with st.expander("Preview extraction data"):
            st.dataframe(df, use_container_width=True)

    except Exception as exc:
        st.error(f"Could not read this file: {exc}")

st.divider()
st.header("1. Manuscript details")

title = st.text_input("Review title")
review_type = st.selectbox(
    "Review type",
    [
        "Systematic review",
        "Systematic review and meta-analysis",
        "Scoping review",
        "Diagnostic test accuracy review",
    ],
)
journal = st.text_input("Target journal (optional)")
citation_style = st.selectbox(
    "Reference style",
    ["Vancouver", "Harvard", "APA", "Other / not decided"],
)

st.header("2. Prepare the manuscript")

st.caption(
    "Enter your verified notes for each section. "
    "The app assembles them into an editable manuscript."
)

sections = {}

with st.expander("Abstract", expanded=True):
    sections["Abstract"] = {}
    for field, prompt in [
        ("Background", "What is the clinical or research problem?"),
        ("Objective", "What was the review objective?"),
        ("Methods", "Databases, eligibility criteria, and synthesis methods"),
        ("Results", "Main verified findings and study counts"),
        ("Conclusion", "What do the findings support?"),
    ]:
        sections["Abstract"][field] = st.text_area(
            field, placeholder=prompt, key=f"abstract_{field}"
        )

section_prompts = {
    "Introduction": (
        "Background, current evidence, knowledge gap, rationale, "
        "and objective. Add reference placeholders such as [REF]."
    ),
    "Methods": (
        "Protocol/registration, eligibility criteria, databases, "
        "search dates, screening, extraction, risk of bias, "
        "and synthesis methods."
    ),
    "Results": (
        "Study selection, study characteristics, risk of bias, "
        "and findings. Verify every number against your data."
    ),
    "Discussion": (
        "Principal findings, comparison with previous research, "
        "possible explanations, clinical implications, and limitations."
    ),
    "Limitations": (
        "Describe limitations of the included evidence and of your review."
    ),
    "Conclusion": (
        "Give a concise conclusion supported by your verified findings. "
        "Avoid claims beyond the evidence."
    ),
    "Declarations": (
        "Funding, conflicts of interest, author contributions, "
        "ethics statement where applicable, and data availability."
    ),
}

for name, prompt in section_prompts.items():
    with st.expander(name, expanded=(name == "Results")):
        sections[name] = st.text_area(
            f"{name} notes and verified findings",
            placeholder=prompt,
            height=160,
            key=f"section_{name}",
        )

st.header("3. Build a Results summary from your data")

auto_results = ""

if df is not None and not df.empty:
    st.write("Choose the columns that describe your included studies.")

    identifier = st.selectbox(
        "Study identifier column",
        ["Not selected"] + list(df.columns),
    )

    available = list(df.columns)

    def choose_column(label):
        return st.selectbox(
            label,
            ["Not selected"] + available,
            key=f"column_{label}",
        )

    year_col = choose_column("Publication year column")
    design_col = choose_column("Study design column")
    sample_col = choose_column("Sample size column")
    findings_col = choose_column("Main findings / outcome column")

    if st.button("Prepare data-based Results draft"):
        parts = []

        if identifier != "Not selected":
            ids = df[identifier].dropna().astype(str).str.strip()
            ids = ids[ids != ""]
            unique_count = ids.nunique()
            parts.append(
                f"The extraction sheet contains {len(df)} rows and "
                f"{unique_count} unique non-empty entries in the "
                f"selected identifier column ({identifier}). "
                "Confirm that these entries represent distinct studies."
            )
        else:
            parts.append(
                f"The extraction sheet contains {len(df)} rows. "
                "A unique study count has not been calculated because "
                "no study identifier was selected."
            )

        for label, col in [
            ("Publication year", year_col),
            ("Study design", design_col),
        ]:
            if col != "Not selected":
                values = df[col].dropna().astype(str).str.strip()
                counts = Counter(v for v in values if v)
                if counts:
                    summary = "; ".join(
                        f"{value} (n={count})"
                        for value, count in counts.most_common(15)
                    )
                    parts.append(
                        f"{label} distribution in the extracted rows: "
                        f"{summary}. Check whether repeated rows belong "
                        "to the same study."
                    )

        if sample_col != "Not selected":
            numeric = pd.to_numeric(df[sample_col], errors="coerce").dropna()
            if not numeric.empty:
                parts.append(
                    f"The selected sample-size column contains "
                    f"{len(numeric)} numeric entries, ranging from "
                    f"{numeric.min():g} to {numeric.max():g}. "
                    "These values are not a pooled participant total; "
                    "check for repeated studies and multiple study arms."
                )

        if findings_col != "Not selected":
            values = df[findings_col].dropna().astype(str).str.strip()
            values = [v for v in values if v]
            if values:
                parts.append(
                    "Extracted findings for verification:\n" +
                    "\n".join(f"- {v}" for v in values[:20])
                )
                if len(values) > 20:
                    parts.append(
                        f"{len(values) - 20} additional entries were "
                        "omitted from this preview."
                    )

        auto_results = "\n\n".join(parts)
        st.session_state["auto_results"] = auto_results

    if st.session_state.get("auto_results"):
        st.text_area(
            "Review and edit the generated Results summary",
            key="edited_results",
            value=st.session_state["auto_results"],
            height=220,
        )
else:
    st.info("Upload an extraction sheet to enable the data-based Results summary.")

st.header("4. Assemble your manuscript")

if st.button("Assemble complete manuscript", type="primary"):
    abstract_text = "\n\n".join(
        f"{field}: {value.strip() or '[Add verified content]'}"
        for field, value in sections["Abstract"].items()
    )

    manuscript_parts = [
        title.strip() or "[Insert manuscript title]",
        f"Review type: {review_type}",
        f"Target journal: {journal.strip() or 'Not specified'}",
        f"Reference style: {citation_style}",
        "",
        "ABSTRACT",
        abstract_text,
    ]

    for name in [
        "Introduction",
        "Methods",
        "Results",
        "Discussion",
        "Limitations",
        "Conclusion",
        "Declarations",
    ]:
        content = sections.get(name, "")
        if name == "Results":
            content = st.session_state.get(
                "edited_results",
                st.session_state.get("auto_results", content),
            ) or content

        manuscript_parts.extend([
            "",
            name.upper(),
            content.strip() or f"[Add verified {name.lower()} content]",
        ])

    manuscript = "\n\n".join(manuscript_parts)
    st.session_state["manuscript"] = manuscript
    st.success("Manuscript assembled. Review it below before exporting.")

if st.session_state.get("manuscript"):
    st.subheader("5. Review and export")
    st.text_area(
        "Complete editable manuscript",
        key="final_manuscript",
        value=st.session_state["manuscript"],
        height=500,
    )

    final_text = st.session_state.get(
        "final_manuscript", st.session_state["manuscript"]
    )

    st.download_button(
        "Download Markdown (.md)",
        data=final_text,
        file_name="systematic_review_manuscript.md",
        mime="text/markdown",
    )

    def make_word_file(text):
        doc = Document()
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            if line in {
                "ABSTRACT", "INTRODUCTION", "METHODS", "RESULTS",
                "DISCUSSION", "LIMITATIONS", "CONCLUSION", "DECLARATIONS"
            }:
                doc.add_heading(line.title(), level=1)
            else:
                doc.add_paragraph(line)
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()

    st.download_button(
        "Download Word (.docx)",
        data=make_word_file(final_text),
        file_name="systematic_review_manuscript.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

st.caption(
    "Important: This is a structured drafting tool, not a substitute "
    "for critical appraisal, statistical analysis, reference verification, "
    "or human scientific review."
)
