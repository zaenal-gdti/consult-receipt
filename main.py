import streamlit as st
from streamlit_option_menu import option_menu
import pandas as pd
from script.mail_merge import MailMerge
from datetime import datetime
import os
import shutil
import time
import glob
import shutil
import concurrent.futures

def main():
    st.set_page_config(page_title="Excel to PDF & PDF Search", layout="wide")
    
    with st.sidebar:
        selected = option_menu(
            menu_title="Navigation", 
            options=["Excel to PDF", "Search PDF"], 
            icons=["file-earmark-spreadsheet", "search"], 
            menu_icon="cast", 
            default_index=0
        )
    
    if selected == "Excel to PDF":
        excel_to_pdf_page()
    elif selected == "Search PDF":
        search_pdf_file()

def run_mail_merge(df, label, file_per_zip = 100):
    #df = pd.read_excel(file)
    if not os.path.exists(f'output/{label}'):
        os.makedirs(f'output/{label}')
    else:
        #shutil.rmtree(f'output/{label}')
        raise Exception("Error: Anda sedang menjalankan mail merge dengan label yang sama, silahkan ganti salah satu label")
    mm = MailMerge(df, label, file_per_zip = file_per_zip) 
    return mm

def move_to_archieved(dir = '.tmp'):
    if not os.path.exists('archieved'):
        os.mkdir('archieved')

    fls = glob.glob(f'{dir}/**/*.pdf', recursive=True)
    for i in fls:
        shutil.copy(i,'archieved')

def excel_to_pdf_page():
    st.title("Excel to PDF Converter")
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.write("Preview of the Excel file:")
        st.dataframe(df)
        if st.button("Convert to PDF"):
            progress_bar = st.progress(0)
            now = datetime.now().strftime('%Y%m%d_%H%M%S')
            zip_output = f'result_{now}'
            mm = run_mail_merge(df, zip_output)

            if os.path.exists(f'.tmp/{mm.label}'):
                shutil.rmtree(f'.tmp/{mm.label}')

            success = []
            errors = []
            k = 0
            for index, row in mm.dataset.iterrows():
                rcp = mm.row_to_pdf(row)
                try:
                    #rcp = self.row_to_pdf(row)
                    success.append(rcp)
                except Exception as e:
                    errors.append(row)
                    print(e)
                progress_bar.progress(int((k + 1) / len(df) * 100))
                k = k + 1
                time.sleep(0.1)
            mm.errors = pd.DataFrame(errors)
            mm.success = pd.DataFrame(success)
            mm.chunk_and_zip()

            move_to_archieved()

            progress_bar.empty()
            st.success("PDF created successfully!")
            with open(f'output/{zip_output}.zip', "rb") as file:
                st.download_button("Download ZIP", file, file_name=f'{zip_output}.zip', mime="application/zip")

def search_pdfs(directory, query):
    matching_files = []
    file_list = [os.path.join(root, file) for root, _, files in os.walk(directory) for file in files if file.lower().endswith(".pdf")]
    file_count = len(file_list)
    
    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = list(executor.map(lambda file: file if query.lower() in os.path.basename(file).lower() else None, file_list))
    
    matching_files = [file for file in results if file is not None]
    return matching_files, file_count

def search_pdf_file():
    st.title("Search for PDF File")
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

if __name__ == "__main__":
    main()
