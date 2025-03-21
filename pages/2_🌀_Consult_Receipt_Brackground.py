import streamlit as st
import requests
import time

# FastAPI Backend URL
FASTAPI_URL = "http://127.0.0.1:8000"

st.set_page_config(page_title="Mail Merge Consult Receipt Background Processing", layout="wide", initial_sidebar_state= "collapsed")
st.write("Upload an Excel file to be processed in the background to create consult receipt.")

# File uploader
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Send file to FastAPI
    with st.spinner("Uploading file..."):
        files = {"file": (uploaded_file.name, uploaded_file.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        response = requests.post(f"{FASTAPI_URL}/upload-excel/", files=files)

    if response.status_code == 200:
        job_id = response.json()["job_id"]
        st.success(f"✅ File uploaded! Job started with ID: `{job_id}`")
       
        # Placeholder to dynamically update job status
        status_placeholder = st.empty()

        # Polling for job status
        status = "⏳ Pending..."
        elapsed_time = "N/A"
        while status not in ["✅ Completed!", "❌ Failed"]:
            time.sleep(2)  # Check status every 2 seconds
            status_response = requests.get(f"{FASTAPI_URL}/job-status/{job_id}")
            status_data = status_response.json()
            status = status_data["status"]
            elapsed_time = status_data["elapsed_time"]

            # 🔄 Refresh job status in the same UI block
            status_placeholder.write(f"📢 **Job Status:** {status}  \n⏳ **Elapsed Time:** {elapsed_time}")

        
        # ✅ Handle ZIP Download Correctly
        if status == "✅ Completed!":
            st.success(f"✅ Job finished in {elapsed_time}! Click below to download your file.")
            download_url = f"{FASTAPI_URL}/download/{job_id}"
            st.markdown(f"[📥 Download Processed ZIP]({download_url})", unsafe_allow_html=True)

    else:
        st.error("❌ File upload failed! Try again.")