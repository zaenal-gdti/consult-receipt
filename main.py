import streamlit as st

st.set_page_config(page_title="PDF Utility App", layout="wide")


st.title("Welcome to the PDF Utility App")

st.write("Use the sidebar to navigate between different tools.")
st.set_page_config(initial_sidebar_state="collapsed")

st.write(st.session_state.processing)

