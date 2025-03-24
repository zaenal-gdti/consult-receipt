import streamlit as st
import requests
import time
from datetime import datetime, timedelta
import pandas as pd


st.set_page_config(page_title="Check Nahsehat Linking", layout="wide", initial_sidebar_state= "collapsed") # Set layout to wide for better display


# Define a global variable to store the token expiration time
bearer_token = None
refresh_token = None
appidbody = st.secretKey['APP_ID']
secretkeybody =  st.secretKey['SECRET_KEY']

# API endpoints 
BASE_URL = st.secretKey['BASE_URL']
AUTH_URL = "/api/v3/auth/GetToken"  
API_URL = "/api/v3/AdmMemberAccess"  

# list of status codes by admMemberAccess
status_codes = {
    "100": "Member Valid",
    "211": "Salah Input Nomor Kartu dan Tanggal Lahir",
    "212": "Nahsehat blocking Nomor Kartu, Silakan hubungi Nahsehat jika hasil Adpass tidak ter blok",
    "213": "Member Tidak Aktif (Admedika)",
    "214": "Polis Member sudah expired",
    "400": "Error Technical (Kontak Group SQA)",
    "401": "Error Technical (Kontak Group SQA)",
    "403": "Error Technical (Kontak Group SQA)",
    "404": "Error Technical (Kontak Admedika)",
    "408": "Error Technical (Kontak Admedika)",
    "500": "Error Technical (Kontak Admedika)",
    "502": "Error Technical (Kontak Admedika)"
}

# Add code here that recurses through the response and prints out tables
# Function to recursively flatten JSON and return as a single dictionary
def flatten_json(data, parent_key='', sep='_'):
    """
    Recursively flattens a nested JSON structure into a flat dictionary.
    """
    items = []
    if isinstance(data, dict):
        for key, value in data.items():
            new_key = f"{parent_key}{sep}{key}" if parent_key else key
            items.extend(flatten_json(value, new_key, sep=sep).items())
    elif isinstance(data, list):
        for i, value in enumerate(data):
            new_key = f"{parent_key}{sep}{i}"
            items.extend(flatten_json(value, new_key, sep=sep).items())
    else:
        items.append((parent_key, data))
    return dict(items)

def authenticate():
    # Placeholder for authentication logic
    # Send request to authentication endpoint
    response = requests.post(BASE_URL + AUTH_URL, data=    {
	"appID" : appidbody,
	"secretKey" : secretkeybody,
	"grantType" : "generate_token"
    })

    if response.status_code == 200:
        # Store the new token and expiration time
        global bearer_token, refresh_token
        bearer_token = response.json()["data"]["accessToken"]
        refresh_token = response.json()["data"]["refreshToken"]
        return True
    else:
        st.error("Authentication failed!" + response.json().get("message"))
        return False

# Streamlit UI
st.title("Look up Admedika Populations")


# Generate Auth Token
# if st.button("Press for auth token"):
#     authenticate()

def call_external_api(card_number, dob):
    if bearer_token is None:
        if not authenticate():
            return None

    # Make API request with card number and date of birth
    headers = {"Authorization": f"Bearer {bearer_token}",
               "Content-Type":"application/json"
    }
    # 8000150102793078
    # 1999-06-10
    data = {"cardno": card_number, "dateofbirth": dob}
    response = requests.post(BASE_URL + API_URL, headers=headers, json=data)
    if response:
        return response.json()  # Return the response data
    else:
        st.error(f"API call failed with status code {response.status_code}")
        return None

# Input fields
card_number = st.text_input("Enter Card Number (Policy Number)")
dob = st.text_input("Enter Date of Birth")
# st.date_input("Enter Date of Birth", min_value=datetime(1900, 1, 1))

# Submit button
if st.button("Retrieve member status eligibility"):
    if not card_number or not dob:
        st.error("Please fill in both fields!")
    else:
        response = call_external_api(card_number, dob)
        if response:
            #Display the API response (for example, display policy details)
            st.success("API Call Successful!")
            
            # st.json(response)
            flattened_data = flatten_json(response)

            # Check if the status code exists in the provided status_codes legend and append the definition
            status_code = str(flattened_data.get("code", ""))  # Get status code as a string
            definition = status_codes.get(status_code, "Unknown Status Code")  # Get definition from the legend

            # Add the definition to the flattened data
            flattened_data['status_definition'] = definition

            # Convert the flattened dictionary into a DataFrame and transpose it to make it vertical
            flattened_df = pd.DataFrame(list(flattened_data.items()), columns=["Field", "Value"])
            st.title("AdMedika Member Information")
            st.table(flattened_df)
        else:
            st.error("Failed to retrieve information from the API.")

# Streamlit UI

# # Display the flattened data as a vertical table
# st.subheader("Flattened JSON Data (Vertical)")
# # st.table(flattened_df)

# # Optionally display the full JSON response for reference
# st.subheader("Full JSON Response")
# st.json(data)
