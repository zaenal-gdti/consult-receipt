from script.mail_merge import MailMerge
import pandas as pd
from datetime import datetime
import os
import shutil
import time
import glob
import streamlit as st
from streamlit_option_menu import option_menu


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

st.set_page_config(page_title="Excel to PDF Converter", page_icon=":shark:")
st.markdown("📄 Excel to PDF Converter")
uploaded_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Preview of the Excel file:")
    st.dataframe(df)
    if st.button("Convert to PDF") :
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
