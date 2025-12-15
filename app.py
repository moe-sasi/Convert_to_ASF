import streamlit as st


def initialize_session_state():
    defaults = {
        "field_mappings": {},
        "config": {},
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def render_sidebar():
    st.sidebar.header("Setup")
    st.sidebar.subheader("ASF template upload")
    st.sidebar.caption("Upload functionality coming soon.")

    st.sidebar.subheader("Loan tape upload")
    st.sidebar.caption("Upload functionality coming soon.")

    st.sidebar.subheader("Fuzzy match threshold")
    st.sidebar.caption("Slider coming soon.")


def render_main_content():
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


def main():
    initialize_session_state()
    render_sidebar()
    render_main_content()


if __name__ == "__main__":
    main()
