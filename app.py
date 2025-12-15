import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from utils import suggest_mappings


def initialize_session_state():
    defaults = {
        "field_mappings": {},
        "config": {},
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    if "field_mappings" not in st.session_state:
        st.session_state["field_mappings"] = {}


def load_asf_template(template_file, sheet_name=None):
    """
    Load ASF template workbook and return (workbook, worksheet).
    If sheet_name is None, use the first sheet.
    """

    template_file.seek(0)
    workbook = load_workbook(template_file)

    worksheet_name = sheet_name or workbook.sheetnames[0]
    worksheet = workbook[worksheet_name]

    return workbook, worksheet


def get_asf_fields(ws, header_row=1):
    """
    Read the header row from the ASF worksheet and return a list of non-empty field names.
    """

    rows = ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True)
    header_cells = next(rows, [])

    return [str(cell).strip() for cell in header_cells if cell is not None and str(cell).strip()]


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

    threshold = st.sidebar.slider(
        "Fuzzy match threshold", min_value=0, max_value=100, value=80
    )

    return asf_template_file, tape_files, threshold


def render_main_content(asf_template_file, tape_files, threshold):
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

    asf_fields = []

    if asf_template_file:
        st.markdown(f"**ASF template uploaded:** {asf_template_file.name}")
        st.caption("ASF template uploaded")

        wb, ws = load_asf_template(asf_template_file)
        asf_fields = get_asf_fields(ws)

        st.markdown("**ASF Fields (header row):**")
        st.write(asf_fields)

    if tape_files:
        for tape_file in tape_files:
            st.markdown(f"### {tape_file.name}")
            dataframe = load_tape_into_dataframe(tape_file)
            st.dataframe(dataframe.head())

            tape_cols = list(dataframe.columns)

            if tape_file.name not in st.session_state["field_mappings"]:
                st.session_state["field_mappings"][tape_file.name] = suggest_mappings(
                    asf_fields, tape_cols, threshold
                )

            st.json(st.session_state["field_mappings"][tape_file.name])


def main():
    initialize_session_state()
    asf_template_file, tape_files, threshold = render_sidebar()
    render_main_content(asf_template_file, tape_files, threshold)


if __name__ == "__main__":
    main()
