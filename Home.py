import streamlit as st
from auth import get_authenticator

st.set_page_config(
    page_title="Odoo Data Cleaner",
    layout="wide"
)

authenticator = get_authenticator()

authenticator.login(location="main")

if st.session_state.get("authentication_status"):

    authenticator.logout(location="sidebar")

    st.title("Odoo Data Cleaning Tools")

    st.markdown("""
    Tools otomatisasi cleaning data Excel Odoo.
    """)

elif st.session_state.get("authentication_status") is False:
    st.error("Username/password incorrect")