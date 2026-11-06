
import io
import re
import pandas as pd
import streamlit as st
from docx import Document

# --------------------------------------------------
# PAGE CONFIGURATION
# --------------------------------------------------

st.set_page_config(
    page_title="Systematic Review Copilot",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --------------------------------------------------
# DARK THEME — MATCH YOUR SECOND SCREENSHOT
# --------------------------------------------------

st.markdown(
    """
    <style>
    :root {
        color-scheme: dark;
    }

    html, body, [data-testid="stAppViewContainer"] {
        background-color: #0e0e0e !important;
        color: #eeeeee !important;
    }

    [data-testid="stHeader"] {
        background-color: #0e0e0e !important;
    }

    [data-testid="stSidebar"] {
        background-color: #171717 !important;
        border-right: 1px solid #303030;
    }

    [data-testid="stMain"] {
        background-color: #0e0e0e !important;
    }

    [data-testid="stMainBlockContainer"] {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1500px;
    }

    h1, h2, h3, h4, p, label,
    [data-testid="stMarkdownContainer"] {
        color: #eeeeee !important;
    }

    .hero {
        background: #202020;
        border: 1px solid #383838;
        border-top: 4px solid #b99b62;
        border-radius: 12px;
        padding: 26px 30px;
        margin-bottom: 26px;
    }

    .hero h1 {
        margin: 0;
        font-size: 2rem;
        color: #ffffff !important;
    }

    .hero p {
        margin: 8px 0 0 0;
        color: #bdbdbd !important;
        font-size: 1rem;
    }

    .section-card {
        background: #1b1b1b;
        border: 1px solid #363636;
        border-radius: 12px;
        padding: 20px;
        margin: 14px 0;
    }

    [data-testid="stTextInput"] input,
    [data-testid="stTextArea"] textarea,
    [data-testid="stSelectbox"] div[data-baseweb="select"],
    [data-testid="stNumberInput"] input {
        background-color: #292929 !important;
        color: #f3f3f3 !important;
        border-color: #494949 !important;
    }

    [data-testid="stFileUploader"] section {
        background-color: #242424 !important;
        border: 1px solid #454545 !important;
        border-radius: 10px;
    }

    [data-testid="stFileUploader"] button {
        background-color: #333333 !important;
        color: #ffffff !important;
        border: 1px solid #555555 !important;
    }

    .stButton > button,
    .stDownloadButton > button {
        background-color: #b99b62 !important;
        color: #111111 !important;
        border: none !important;
        border-radius: 7px !important;
        font-weight: 600 !important;
        min-height: 42px;
    }

    .stButton > button:hover,
    .stDownloadButton > button:hover {
        background-color: #d0b57d !important;
        color: #111111 !important;
    }

    [data-testid="stMetric"] {
        background: #202020;
        border: 1px solid #3b3b3b;
        border-radius: 10px;
        padding: 16px;
    }

    [data-testid="stMetricValue"] {
        color: #d5bd8a !important;
    }

    [data-testid="stDataFrame"] {
        border: 1px solid #3a3a3a;
        border-radius: 8px;
    }

    [data-testid="stExpander"] {
        background: #1b1b1b;
        border: 1px solid #383838;
        border-radius: 8px;
    }

    hr {
        border-color: #383838 !important;
    }

    .muted {
        color: #aaaaaa;
        font-size: 0.9rem;
    }

    footer {
        visibility: hidden;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------
# MANUSCRIPT SECTIONS
# --------------------------------------------------

SECTIONS = [
    "Abstract",
    "Introduction",
    "Materials and Methods",
    "Results",
    "Discussion",
    "Limitations",
    "Conclusion",
    "Declarations",
    "References",
]

SECTION_GUIDANCE = {
    "Abstract": (
        "Provide the background, objective, methods, principal "
        "results, and conclusion. Include verified numerical results."
    ),
    "Introduction": (
        "Describe the background, clinical or scientific problem, "
        "existing evidence, knowledge gap, and review objective."
    ),
    "Materials and Methods": (
        "Describe the protocol, eligibility criteria, databases, "
        "search dates, screening, extraction, risk-of-bias assessment, "
        "and synthesis methods."
    ),
    "Results": (
        "Report PRISMA counts, included-study characteristics, "
        "risk of bias, individual outcomes, and synthesis results."
    ),
    "Discussion": (
        "Interpret the main findings, compare them with previous "
        "research, discuss implications, and explain uncertainties."
    ),
    "Limitations": (
        "Describe limitations of the evidence, review methods, "
        "and applicability of the findings."
    ),
    "Conclusion": (
        "Summarise the principal findings and their implications "
        "without making claims unsupported by the evidence."
    ),
    "Declarations": (
        "Add funding, conflicts of interest, ethics, data availability, "
        "author contributions, and other applicable declarations."
    ),
    "References": (
        "Insert the verified reference list in the required journal style."
    ),
}

# --------------------------------------------------
# SIDEBAR
# --------------------------------------------------

with st.sidebar:
    st.markdown("## 📚 SR Copilot")
    st.caption("Evidence synthesis workspace")
    st.divider()

    page = st.radio(
        "Workspace",
        [
            "Dashboard",
            "Review Setup",
            "PRISMA Calculator",
            "Study Screening",
            "Manuscript Studio",
        ],
        index=4,
    )

    st.divider()
    st.caption("Manuscript Studio")
    st.markdown(
        """
        - Structured manuscript sections
        - Extraction-sheet inspection
        - Editable drafts
        - Markdown export
        - Word document export
        """
    )
    st.caption("Template-based writing · No AI API required")

# --------------------------------------------------
# HEADER
# --------------------------------------------------

st.markdown(
    """
    <div class="hero">
        <h1>📚 Systematic Review Copilot</h1>
        <p>Your workspace for evidence synthesis and manuscript preparation</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------
# SESSION STATE
# --------------------------------------------------

if "manuscript" not in st.session_state:
    st.session_state.manuscript = {}

if "draft" not in st.session_state:
    st.session_state.draft = ""

if "review_title" not in st.session_state:
    st.session_state.review_title = ""

# --------------------------------------------------
# DOCUMENT EXPORT FUNCTIONS
# --------------------------------------------------

def make_markdown(title, details, sections):
    lines = [f"# {title or 'Systematic Review Manuscript'}", ""]

    if details.get("journal"):
        lines.extend([f"**Target journal:** {details['journal']}", ""])

    if details.get("review_type"):
        lines.extend([f"**Review type:** {details['review_type']}", ""])

    lines.extend(["---", ""])

    for section in SECTIONS:
        content = sections.get(section, "").strip()

        lines.extend([f"## {section}", ""])

        if content:
            lines.extend([content, ""])
        else:
            lines.extend([
                f"[Complete this section: {SECTION_GUIDANCE[section]}]",
                "",
            ])

    lines.extend([
        "---",
        "",
        "*Draft prepared from user-provided research information. "
        "Verify all findings, citations, and reporting requirements "
        "before submission.*",
    ])

    return "\n".join(lines)


def make_word(markdown_text):
    doc = Document()

    doc.add_heading("Systematic Review Manuscript", level=0)

    for line in markdown_text.splitlines():
        text = line.strip()

        if not text or text == "---":
            continue

        if text.startswith("# "):
            doc.add_heading(text[2:], level=0)

        elif text.startswith("## "):
            doc.add_heading(text[3:], level=1)

        elif text.startswith("### "):
            doc.add_heading(text[4:], level=2)

        elif text.startswith("**") and text.endswith("**"):
            doc.add_paragraph(text.strip("*"))

        elif text.startswith("*") and text.endswith("*"):
            doc.add_paragraph(text.strip("*"))

        else:
            doc.add_paragraph(text)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# --------------------------------------------------
# DASHBOARD
# --------------------------------------------------

if page == "Dashboard":
    st.title("Dashboard")
    st.write(
        "Organise your review, inspect your extracted evidence, "
        "and prepare a manuscript using verified research information."
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Manuscript sections", len(SECTIONS))

    with col2:
        st.metric("Export formats", 2)

    with col3:
        st.metric("AI API required", "No")

    st.markdown("### Getting started")
    st.markdown(
        """
        1. Open **Review Setup** and enter your review details.
        2. Use **PRISMA Calculator** to check screening counts.
        3. Use **Study Screening** to organise study decisions.
        4. Open **Manuscript Studio** and enter your verified findings.
        5. Generate your structured draft and export it to Word or Markdown.
        """
    )

# --------------------------------------------------
# REVIEW SETUP
# --------------------------------------------------

elif page == "Review Setup":
    st.title("Review Setup")

    with st.form("review_setup_form"):
        title = st.text_input(
            "Review title",
            value=st.session_state.review_title,
            placeholder="Enter your systematic review title",
        )

        question = st.text_area(
            "Research question",
            placeholder="Enter your PICO, PEO, or review question",
        )

        eligibility = st.text_area(
            "Inclusion and exclusion criteria",
            placeholder="Define your eligibility criteria",
        )

        review_type = st.selectbox(
            "Review type",
            [
                "Systematic review",
                "Systematic review and meta-analysis",
                "Scoping review",
                "Diagnostic test accuracy review",
                "Other",
            ],
        )

        journal = st.text_input(
            "Target journal (optional)"
        )

        reference_style = st.selectbox(
            "Reference style",
            ["Vancouver", "Harvard", "APA", "AMA", "Other"],
        )

        submitted = st.form_submit_button("Save review details")

        if submitted:
            st.session_state.review_title = title
            st.session_state.review_details = {
                "title": title,
                "question": question,
                "eligibility": eligibility,
                "review_type": review_type,
                "journal": journal,
                "reference_style": reference_style,
            }
            st.success("Review details saved for this session.")

# --------------------------------------------------
# PRISMA CALCULATOR
# --------------------------------------------------

elif page == "PRISMA Calculator":
    st.title("PRISMA Calculator")
    st.write("Enter your verified study-selection counts.")

    with st.form("prisma_form"):
        identified = st.number_input(
            "Records identified", min_value=0, step=1
        )
        duplicates = st.number_input(
            "Duplicate records removed", min_value=0, step=1
        )
        screened = st.number_input(
            "Records screened", min_value=0, step=1
        )
        title_excluded = st.number_input(
            "Records excluded at title/abstract screening",
            min_value=0,
            step=1,
        )
        full_text = st.number_input(
            "Full-text reports assessed", min_value=0, step=1
        )
        full_text_excluded = st.number_input(
            "Full-text reports excluded", min_value=0, step=1
        )
        included = st.number_input(
            "Studies included in review", min_value=0, step=1
        )

        prisma_submit = st.form_submit_button("Check counts")

    if prisma_submit:
        expected_screened = identified - duplicates
        expected_full_text = screened - title_excluded
        expected_included = full_text - full_text_excluded

        st.markdown("### Count checks")

        checks = [
            (
                "Records screened",
                screened,
                expected_screened,
            ),
            (
                "Full-text reports assessed",
                full_text,
                expected_full_text,
            ),
            (
                "Studies included",
                included,
                expected_included,
            ),
        ]

        for label, actual, expected in checks:
            if actual == expected:
                st.success(f"{label}: count is internally consistent.")
            else:
                st.warning(
                    f"{label}: entered {actual}; arithmetic check "
                    f"suggests {expected}. Verify the definitions and counts."
                )

        st.caption(
            "These are arithmetic checks, not a validated PRISMA flow diagram. "
            "Account separately for reports and studies where applicable."
        )

# --------------------------------------------------
# STUDY SCREENING
# --------------------------------------------------

elif page == "Study Screening":
    st.title("Study Screening")
    st.write(
        "Upload a screening spreadsheet to inspect study decisions. "
        "This module does not independently assess eligibility."
    )

    screening_file = st.file_uploader(
        "Upload screening data",
        type=["xlsx", "csv"],
        key="screening_upload",
    )

    if screening_file:
        try:
            if screening_file.name.endswith(".csv"):
                screening_df = pd.read_csv(screening_file)
            else:
                screening_df = pd.read_excel(screening_file)

            st.success("Screening file loaded.")
            st.metric("Rows", len(screening_df))
            st.dataframe(screening_df, use_container_width=True)

        except Exception as e:
            st.error(f"Could not read the file: {e}")

# --------------------------------------------------
# MANUSCRIPT STUDIO
# --------------------------------------------------

elif page == "Manuscript Studio":
    st.title("Manuscript Studio")

    st.markdown(
        """
        <p style="color:#aaaaaa;">
        Create a structured manuscript from your verified research information.
        This version uses templates, not an AI writing API, and does not invent
        references or study findings.
        </p>
        """,
        unsafe_allow_html=True,
    )

    # Upload extraction sheet
    st.markdown("### Upload your study extraction sheet")

    uploaded_file = st.file_uploader(
        "Upload an Excel or CSV extraction sheet",
        type=["xlsx", "csv"],
        key="manuscript_upload",
    )

    df = None

    if uploaded_file:
        try:
            if uploaded_file.name.lower().endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.success("Extraction sheet loaded successfully.")

            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric("Rows", len(df))

            with col2:
                st.metric("Columns", len(df.columns))

            with col3:
                st.metric("Missing cells", int(df.isna().sum().sum()))

            with st.expander("Preview extraction data"):
                st.dataframe(df, use_container_width=True)

            with st.expander("Column names"):
                st.write(", ".join(map(str, df.columns)))

        except Exception as e:
            st.error(f"Could not read the file: {e}")

    st.divider()

    # Manuscript details
    st.markdown("### 1. Manuscript details")

    details = st.session_state.get("review_details", {})

    title = st.text_input(
        "Manuscript title",
        value=details.get("title", st.session_state.review_title),
        key="manuscript_title",
    )

    col1, col2 = st.columns(2)

    with col1:
        review_type = st.selectbox(
            "Review type",
            [
                "Systematic review",
                "Systematic review and meta-analysis",
                "Scoping review",
                "Diagnostic test accuracy review",
                "Other",
            ],
            index=(
                [
                    "Systematic review",
                    "Systematic review and meta-analysis",
                    "Scoping review",
                    "Diagnostic test accuracy review",
                    "Other",
                ].index(details["review_type"])
                if details.get("review_type") in [
                    "Systematic review",
                    "Systematic review and meta-analysis",
                    "Scoping review",
                    "Diagnostic test accuracy review",
                    "Other",
                ]
                else 0
            ),
        )

    with col2:
        journal = st.text_input(
            "Target journal",
            value=details.get("journal", ""),
        )

    reference_style = st.selectbox(
        "Reference style",
        ["Vancouver", "Harvard", "APA", "AMA", "Other"],
        index=(
            ["Vancouver", "Harvard", "APA", "AMA", "Other"].index(
                details["reference_style"]
            )
            if details.get("reference_style") in [
                "Vancouver", "Harvard", "APA", "AMA", "Other"
            ]
            else 0
        ),
    )

    research_question = st.text_area(
        "Research question and objective",
        value=details.get("question", ""),
        placeholder="Enter the main question and review objective.",
        height=100,
    )

    eligibility = st.text_area(
        "Eligibility criteria",
        value=details.get("eligibility", ""),
        placeholder="Enter your inclusion and exclusion criteria.",
        height=100,
    )

    # Section-specific research notes
    st.markdown("### 2. Add your verified research information")
    st.caption(
        "Enter your actual findings in each section. Empty sections will "
        "receive prompts rather than fabricated manuscript content."
    )

    section_notes = {}

    for section in SECTIONS:
        with st.expander(section, expanded=(section == "Abstract")):
            st.caption(SECTION_GUIDANCE[section])

            existing = st.session_state.manuscript.get(section, "")

            section_notes[section] = st.text_area(
                f"Notes for {section}",
                value=existing,
                placeholder=(
                    f"Paste your verified information for {section} here..."
                ),
                height=180,
                key=f"notes_{section}",
                label_visibility="collapsed",
            )

    # Optional extraction-sheet summary
    st.markdown("### 3. Extraction-sheet summary")

    study_column = "None"

    if df is not None:
        study_column = st.selectbox(
            "Select the study identifier column",
            ["None"] + list(map(str, df.columns)),
        )

        st.caption(
            "Use a column that identifies studies. Multiple rows may belong "
            "to one study, so row counts and unique-study counts can differ."
        )

    # Generate draft
    if st.button("Prepare complete manuscript draft", type="primary"):
        manuscript_sections = {
            section: section_notes.get(section, "").strip()
            for section in SECTIONS
        }

        # Use setup details only when the section has not already been filled.
        if not manuscript_sections["Introduction"] and research_question:
            manuscript_sections["Introduction"] = (
                f"Review question and objective: {research_question}\n\n"
                + (
                    f"Eligibility criteria: {eligibility}"
                    if eligibility
                    else "[Add the background, knowledge gap, and rationale.]"
                )
            )

        # Include descriptive extraction-sheet information without
        # treating each row as a unique study.
        if df is not None:
            rows = len(df)
            columns = len(df.columns)

            if study_column != "None":
                unique_studies = int(df[study_column].nunique(dropna=True))
                data_summary = (
                    f"Extraction sheet contains {rows} rows and {columns} "
                    f"columns. The selected identifier column contains "
                    f"{unique_studies} unique non-missing values. Confirm "
                    f"that each value represents a distinct study and that "
                    f"the extraction sheet has been correctly deduplicated "
                    f"before reporting study counts."
                )
            else:
                data_summary = (
                    f"The extraction sheet contains {rows} rows and "
                    f"{columns} columns. A study identifier was not selected, "
                    f"so the number of unique studies has not been calculated."
                )

            if manuscript_sections["Results"]:
                manuscript_sections["Results"] += "\n\n" + data_summary
            else:
                manuscript_sections["Results"] = data_summary

        st.session_state.manuscript = manuscript_sections

        draft = make_markdown(
            title=title,
            details={
                "journal": journal,
                "review_type": review_type,
                "reference_style": reference_style,
            },
            sections=manuscript_sections,
        )

        st.session_state.draft = draft
        st.session_state.review_title = title

        st.success(
            "Structured draft prepared. Review each section and verify "
            "all findings, numerical results, and references."
        )

    # Draft preview and downloads
    if st.session_state.draft:
        st.divider()
        st.markdown("### 4. Review and export your manuscript")

        edited_draft = st.text_area(
            "Editable complete manuscript",
            value=st.session_state.draft,
            height=500,
            key="editable_complete_draft",
        )

        # Update saved draft to match the edited text
        st.session_state.draft = edited_draft

        st.download_button(
            "Download Markdown",
            data=edited_draft,
            file_name="systematic_review_manuscript.md",
            mime="text/markdown",
        )

        try:
            word_file = make_word(edited_draft)

            st.download_button(
                "Download Word document",
                data=word_file,
                file_name="systematic_review_manuscript.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
            )

        except Exception:
            st.warning(
                "Word export requires python-docx. Add python-docx to "
                "requirements.txt and redeploy the app."
            )

        st.caption(
            "This is a structured, template-based draft. It does not "
            "independently write a scientific argument, verify citations, "
            "perform meta-analysis, or validate the review methodology."
        )

# --------------------------------------------------
# FOOTER
# --------------------------------------------------

st.divider()

st.markdown(
    """
    <p style="text-align:center;color:#888888;font-size:0.85rem;">
    Systematic Review Copilot · Research assistance with human verification
    </p>
    """,
    unsafe_allow_html=True,
)
