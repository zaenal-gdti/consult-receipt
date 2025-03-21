import streamlit as st
import pandas as pd
import duckdb
from st_aggrid import AgGrid, GridOptionsBuilder

st.set_page_config(layout="wide", initial_sidebar_state= "collapsed") # Set layout to wide for better display

st.title("[Prototype] - Input Claim Detail Consult")

# Connect to DuckDB database
def get_connection():
    conn = duckdb.connect(database="data.duckdb", read_only=False)
    return conn

def load_data_from_db():
    try:
        conn = get_connection()
        df = conn.execute("SELECT * FROM consult_df").fetchdf()
        conn.close()
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()  # Return empty DataFrame instead of removing table

def save_data_to_db(df):
    try:
        conn = get_connection()
        conn.execute("CREATE TABLE IF NOT EXISTS consult_df AS SELECT * FROM df")
        conn.execute("DELETE FROM consult_df")  # Clear table before inserting new data
        conn.execute("INSERT INTO consult_df SELECT * FROM df")
        conn.close()
    except Exception as e:
        st.error(f"Error saving data: {e}")

def highlight_missing(s):
    return ['background-color: #ffcccc' if pd.isnull(v) or v == '' else '' for v in s]

st.write("### Data Table")
df = load_data_from_db()

if not df.empty:
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(editable=True)
    grid_options = gb.build()

    grid_response = AgGrid(df, gridOptions=grid_options, enable_enterprise_modules=False, update_mode='MANUAL', fit_columns_on_grid_load=True, height=600, width='100%')

    df = grid_response['data']
    save_data_to_db(df)

    st.write("### Updated Data Table")
    st.dataframe(df.style.apply(highlight_missing, axis=1), width=1500)  # Highlight missing cells
else:
    st.warning("No data available.")
