import streamlit as st

st.set_page_config(
    page_title="GL Report",
    layout="wide",
)

st.title("General Ledger Report")

st.sidebar.success("Select a report type.")

st.write("Please select a report type at the left sidebar.")
