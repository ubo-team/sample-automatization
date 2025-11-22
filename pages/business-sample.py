import streamlit as st

st.markdown("""
    <div style='width: 100%; padding: 20px 30px; background: #ffffff;
                border-bottom: 1px solid #e6e6e6; display: flex;
                justify-content: space-between; align-items: center;'>
        <a href="/" style='font-size: 18px; font-weight: 600; color: #344b77;
                text-decoration: none;'>← Faqja kryesore</a>
    </div>
"""
            , unsafe_allow_html=True)

st.markdown("""
<style>
/* Remove sidebar */
section[data-testid="stSidebar"] {
    display: none !important;
}

</style>
""", unsafe_allow_html=True)


# =========================
# CONFIG
# =========================

st.set_page_config(
    page_title="Dizajnimi i Mostrës së Bizneseve",
    layout="wide"
)


st.title("Dizajnimi i Mostrës së Bizneseve")

st.markdown("""
Kjo faqe është duke u zhvilluar...""")