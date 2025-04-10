import streamlit as st

# Dummy user data (for demo purposes)
USER_CREDENTIALS = {
    "alice": "password123",
    "bob": "secure456"
}

def login():
    st.title("🔐 Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    login_btn = st.button("Login")

    if login_btn:
        if username in USER_CREDENTIALS and USER_CREDENTIALS[username] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success(f"Welcome, {username}!")
            st.rerun()
        else:
            st.error("Invalid username or password")

def logout():
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.experimental_rerun()

def main():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if st.session_state.logged_in:
        st.sidebar.success(f"Logged in as {st.session_state.username}")
        if st.sidebar.button("Logout"):
            logout()
        st.write("Use the sidebar to navigate between pages.")
    else:
        login()

if __name__ == "__main__":
    main()


# # Login function
# def login():
#     st.sidebar.title("Login")
#     username = st.sidebar.text_input("Username")
#     password = st.sidebar.text_input("Password", type="password")
#     login_button = st.sidebar.button("Login")

#     if login_button:
#         if username in USER_CREDENTIALS and USER_CREDENTIALS[username] == password:
#             st.session_state.logged_in = True
#             st.session_state.username = username
#         else:
#             st.sidebar.error("Invalid username or password")


# # App flow
# if "logged_in" not in st.session_state:
#     st.session_state.logged_in = False

# if st.session_state.logged_in:
#     main_app()
# else:
#     login()