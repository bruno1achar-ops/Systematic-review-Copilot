
import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="SR Manuscript Copilot",
    page_icon="📚",
    layout="wide"
)

st.title("📚 SR Manuscript Copilot")
st.write(
    "Turn your systematic review extraction data "
    "into a structured manuscript draft."
)

st.subheader("1. Upload your extraction sheet")

uploaded_file = st.file_uploader(
    "Upload an Excel file",
    type=["xlsx"]
)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        st.success("File uploaded successfully!")
        st.write("Preview of your data")
        st.dataframe(df)

        st.subheader("2. Study characteristics")

        st.write(f"Number of rows: {len(df)}")
        st.write(f"Number of columns: {len(df.columns)}")

        st.write("Column names:")
        st.write(", ".join(map(str, df.columns)))

        st.subheader("3. Results draft")

        study_column = st.selectbox(
            "Select the study identifier column",
            ["None"] + list(df.columns.astype(str))
        )

        if st.button("Generate draft"):
            if study_column == "None":
                st.warning(
                    "Select the study identifier column first."
                )
            else:
                n = df[study_column].nunique()
                draft = (
                    f"The extraction sheet contains {len(df)} "
                    f"rows representing {n} unique entries "
                    f"in the selected study identifier column. "
                    "The extracted variables are presented in "
                    "the accompanying data table. Confirm that "
                    "each entry represents a distinct study "
                    "before reporting these counts in the manuscript."
                )

                st.text_area(
                    "Editable draft",
                    value=draft,
                    height=180
                )

                st.download_button(
                    "Download draft",
                    data=draft,
                    file_name="results_draft.txt"
                )

    except Exception as e:
        st.error(f"Could not read the file: {e}")
else:
    st.info("Upload your study extraction Excel sheet to begin.")
