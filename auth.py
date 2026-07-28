import streamlit as st
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader


def get_authenticator():
    if "authenticator" not in st.session_state:

        with open("config.yaml") as file:
            config = yaml.load(file, Loader=SafeLoader)

        st.session_state.authenticator = stauth.Authenticate(
            config["credentials"],
            config["cookie"]["name"],
            config["cookie"]["key"],
            config["cookie"]["expiry_days"],
        )

    return st.session_state.authenticator


def require_login():
    authenticator = get_authenticator()

    # restore dari cookie
    authenticator.login(location="unrendered")

    if st.session_state.get("authentication_status"):
        return True

    elif st.session_state.get("authentication_status") is False:
        st.error("Username/password incorrect")
        st.stop()

    else:
        st.warning("Please login first")
        st.switch_page("Home.py")
        st.stop()