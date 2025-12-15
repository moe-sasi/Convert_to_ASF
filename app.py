import pandas as pd
import streamlit as st


def initialize_session_state():
    defaults = {
        "field_mappings": {},
        "config": {},
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def load_tape_into_dataframe(file):
    filename = file.name.lower()

    if filename.endswith(".csv"):
        return pd.read_csv(file)

    if filename.endswith(".xlsx"):
        return pd.read_excel(file)

    raise ValueError("Unsupported file type. Please upload a .csv or .xlsx file.")


def render_sidebar():
    st.sidebar.header("Setup")
    st.sidebar.subheader("ASF template upload")
    asf_template_file = st.sidebar.file_uploader(
        "Upload ASF template",
        type=["xlsx"],
        key="asf_template_file",
    )

    st.sidebar.subheader("Loan tape upload")
    tape_files = st.sidebar.file_uploader(
        "Upload loan tape files",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key="tape_files",
    )

    st.sidebar.subheader("Fuzzy match threshold")
    st.sidebar.caption("Slider coming soon.")

    return asf_template_file, tape_files


def render_main_content(asf_template_file, tape_files):
    st.title("ASF Loan Tape Mapper")
    st.subheader("Workflow")
    st.markdown(
        """
        1. Upload ASF template
        2. Upload loan tapes
        3. Review mapping
        4. Download ASF output
        """
    )

    if not asf_template_file and not tape_files:
        st.info("Upload ASF template and loan tape files to begin.")
        return

    if asf_template_file:
        st.markdown(f"**ASF template uploaded:** {asf_template_file.name}")
        st.caption("ASF template uploaded")

    if tape_files:
        for tape_file in tape_files:
            st.markdown(f"### {tape_file.name}")
            dataframe = load_tape_into_dataframe(tape_file)
            st.dataframe(dataframe.head())


def main():
    initialize_session_state()
    asf_template_file, tape_files = render_sidebar()
    render_main_content(asf_template_file, tape_files)


if __name__ == "__main__":
    main()
