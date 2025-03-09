import streamlit as st
from streamlit_option_menu import option_menu
import concurrent.futures
import os

def search_pdfs(directory, query):
    matching_files = []
    file_list = [os.path.join(root, file) for root, _, files in os.walk(directory) for file in files if file.lower().endswith(".pdf")]
    file_count = len(file_list)
    
    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = list(executor.map(lambda file: file if query.lower() in os.path.basename(file).lower() else None, file_list))
    
    matching_files = [file for file in results if file is not None]
    return matching_files, file_count


st.set_page_config(page_title="Search PDF Files", page_icon=":wolf:")
st.title("🔍 Search PDF Files")

directory = 'archieved'
query = st.text_input("Enter PDF filename to search for:")
if st.button("Search") and directory and query:
    progress_bar = st.progress(0)
    matching_files, file_count = search_pdfs(directory, query)
    progress_bar.progress(100)
    progress_bar.empty()
    if matching_files:
        st.write("Found PDF files:")
        for file in matching_files:
            with open(file, "rb") as pdf_file:
                st.download_button(label=f"Download {os.path.basename(file)}", data=pdf_file, file_name=os.path.basename(file), mime="application/pdf")
    else:
        st.write("No matching PDF files found.")
    