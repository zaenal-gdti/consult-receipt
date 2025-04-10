import streamlit as st
import requests
import time
import os 

FASTAPI_URL = "http://127.0.0.1:8000"

if "logged_in" not in st.session_state or not st.session_state.logged_in:
    st.warning("Please log in from the main page.")
    st.stop()

st.title("📋 Job Monitor & Downloads")



# Display all jobs with optional download buttons
st.header("📡 All Jobs")
jobs_placeholder = st.empty()

def fetch_jobs():
    response = requests.get(f"{FASTAPI_URL}/all-jobs/")
    if response.status_code == 200:
        return response.json()["jobs"]
    return []

# Refresh job list every 5 seconds
while True:
    time.sleep(10)
    jobs = fetch_jobs()

    with jobs_placeholder.container():
        if not jobs:
            st.write("✅ No jobs yet.")
        else:
            for job in jobs:
                cols = st.columns([3, 3, 2, 2])
                cols[0].markdown(f"**🆔 Job ID:** `{job['job_id']}`")
                cols[1].markdown(f"**📢 Status:** {job['status']}")
                cols[2].markdown(f"**⏳ Time:** {job['elapsed_time']}")

                if "Completed" in job["status"]:
                    download_url = f"{FASTAPI_URL}/download/{job['job_id']}"
                    cols[3].markdown(f"[⬇️ Download ZIP]({download_url})", unsafe_allow_html=True)
                else:
                    cols[3].write("⏳ Processing")