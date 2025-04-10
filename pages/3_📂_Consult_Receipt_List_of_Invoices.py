import streamlit as st
import os
import datetime
# Title
st.title("📂 Consult Receipt List of Created Invoices")

if "logged_in" not in st.session_state or not st.session_state.logged_in:
    st.warning("Please log in from the main page.")
    st.stop()


# Set directory
directory = "output"  # Change to your folder path

# Get list of files with creation time
if os.path.isdir(directory):
    files = [
        {
            "name": f,
            "path": os.path.join(directory, f),
            "ctime": os.path.getctime(os.path.join(directory, f))
        }
        for f in os.listdir(directory)
        if os.path.isfile(os.path.join(directory, f))
    ]

    # Sort by creation time (descending: newest first)
    files.sort(key=lambda x: x["ctime"], reverse=True)
else:
    st.error("Directory not found.")
    files = []

# Display as a table
if files:
    st.write("### Files (Newest First)")
    for i, file in enumerate(files):
        filename = file["name"]
        file_path = file["path"]
        created_time = datetime.datetime.fromtimestamp(file["ctime"]).strftime('%Y-%m-%d %H:%M:%S')

        col1, col2, col3, col4 = st.columns([1, 4, 3, 2])
        col1.write(i + 1)
        col2.write(filename)
        col3.write(created_time)

        with open(file_path, "rb") as f:
            file_bytes = f.read()

        col4.download_button(
            label="Download",
            data=file_bytes,
            file_name=filename,
            mime="application/octet-stream",
            key=f"download_{i}"
        )
else:
    st.info("No files found in the directory.")