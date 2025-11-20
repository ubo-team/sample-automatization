import streamlit as st

st.set_page_config(
    page_title="Dizajnimi i mostrës",
    page_icon="📊",
    layout="wide"
)

def load_svg(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

# =====================================================
# CUSTOM CSS
# =====================================================
st.markdown("""
<style>
            
.logo svg {
    width: 120px;
    height: auto;
}
            
.icon svg {
    width: 60px;
    height: auto;
}

.title {
    font-size: 44px;
    font-weight: 700;
    margin-top: 20px;
    margin-bottom: 30px;
}     

.stButton > button {
    width: 500px !important;
    display: flex !important;
    justify-content: space-between !important;
    align-items: center !important;

    border-radius: 14px !important;
    border: 1px solid #e3e6eb !important;
    background-color: #ffffff !important;

    font-size: 20px !important;
    font-weight: 600 !important;
    color: #333333 !important;

    padding: 18px 22px !important;
    text-align: left !important;

    transition: all 0.25s ease !important;
}

 /* ===== BUTTON HOVER ===== */
.stButton > button:hover {
    background-color: #eef4ff !important;
    border-color: #4c8bf5 !important;
    color: #1a73e8 !important;
    box-shadow: 0px 4px 10px rgba(76,139,245,0.12) !important;
}

/* Hide ENTIRE sidebar */
section[data-testid="stSidebar"] {
    display: none !important;
}

/* Expand main content to full width */
div[data-testid="stAppViewContainer"] > .main {
    margin-left: 0 !important;
    padding-left: 0 !important;
    padding-right: 0 !important;
}

</style>
""", unsafe_allow_html=True)

# =====================================================
# CUSTOM HTML BUTTON
# Uses JS → query param → switch_page
# =====================================================
icon_1 = load_svg("images/nation.svg")
icon_2 = load_svg("images/municipality.svg")
icon_3 = load_svg("images/business.svg")


def modern_button(label, page, icon):
    if st.button(f"{icon}{label}"):
        st.switch_page(page)

def modern_button(label, page, icon):

    icon_col, btn_col = st.columns([0.12, 0.88])

    with icon_col:
        st.markdown(f"<div class='icon'>{icon}</div>", unsafe_allow_html=True)

    with btn_col:
        if st.button(label):
            st.switch_page(page)


# =====================================================
# LAYOUT
# =====================================================
left, right = st.columns([1.5, 1])

with left:

    # Logo
    with open("images/UBO logo.svg", "r", encoding="utf-8") as f:
        svg_logo = f.read()
    st.markdown(f"<div class='logo'>{svg_logo}</div>", unsafe_allow_html=True)

    # Title
    st.markdown("<div class='title'>Dizajnimi i mostrës</div>", unsafe_allow_html=True)

    # Buttons (fully styled)
    modern_button("Mostra nacionale", "pages/national-sample.py", icon_1)
    modern_button("Mostra komunale", "pages/2_mostra_komunale.py", icon_2)
    modern_button("Mostra për biznese", "pages/3_mostra_biznese.py", icon_3)

with right:
    st.image("images/home.png", use_container_width=True)
