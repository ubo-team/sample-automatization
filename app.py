import streamlit as st

st.set_page_config(
    page_title="Dizajnimi i mostrës",
    page_icon="images/page-icon.png",
    layout="wide"
)

# =====================================================
# CSS
# =====================================================
st.markdown("""
<style>

.logo svg {
    width: 100px;
    height: auto;
}

.title {
    font-size: 44px;
    font-weight: 700;
    margin-top: 20px;
    margin-bottom: 40px;
    text-align: center;
    position: absolute;
    left: 50%;
    top: 0;
    transform: translate(-50%, -125%);
}

.stHorizontalBlock {
    gap: 1.5rem !important;
}
            
.stVerticalBlock {
    align-items: center;
}
            
.stVerticalBlock > div:nth-child(2):has(button) {
    width: 80% !important
}

/* Center each column content */
.col-container {
    text-align: center;
    padding: 20px;
}

/* Image styling */
.col-image img {
    width: 140px;
    height: auto;
    margin-bottom: 15px;
}

/* Title */
.col-title {
    font-size: 26px;
    font-weight: 700;
    margin-bottom: 10px;
    text-align: center;
}

/* Description */
.col-text {
    font-size: 16px;
    color: #555;
    margin-bottom: 20px;
    text-align: center;
}

/* Buttons */
.stButton > button {
    padding: 10px 18px !important;
    font-size: 16px !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    border: 2px solid #344b77 !important;
    color: #344b77 !important;
    background: #fff !important;
    transition: 0.25s ease !important;
}
            
.stButton > button:hover {
    background: #344b77 !important;
    color: white !important;
}

/* Remove sidebar */
section[data-testid="stSidebar"] {
    display: none !important;
}

div[data-testid="stAppViewContainer"] > .main {
    margin-left: 0 !important;
}

</style>
""", unsafe_allow_html=True)

# =====================================================
# COLUMN COMPONENT WITHOUT CARDS
# =====================================================
def menu_column(title, description, page, image_path, key):
    st.markdown("<div class='col-container'>", unsafe_allow_html=True)

    # Image
    st.image(image_path, use_container_width=True)

    # Title
    st.markdown(f"<div class='col-title'>{title}</div>", unsafe_allow_html=True)

    # Description
    st.markdown(f"<div class='col-text'>{description}</div>", unsafe_allow_html=True)

    # Button
    if st.button("Fillo gjenerimin", key=key):
        st.switch_page(page)

    st.markdown("</div>", unsafe_allow_html=True)


# =====================================================
# PAGE LAYOUT
# =====================================================

# Logo
with open("images/UBO Logo.svg", "r", encoding="utf-8") as f:
    svg_logo = f.read()
st.markdown(f"<div class='logo' style='text-align:left'>{svg_logo}</div>", unsafe_allow_html=True)

# Title
st.markdown("<div class='title'>Dizajnimi i mostrës</div>", unsafe_allow_html=True)

# 3 Columns
col1, col2, col3 = st.columns(3)

with col1:
    menu_column(
        "Mostra nacionale",
        "Gjeneroni ndarjen e mostrës në nivel nacional sipas komunës/regjionit, etnisë dhe vendbanimit. Lejon aplikimin e filterëve demografikë dhe oversampling.",
        "pages/national-sample.py",
        "images/nation.png",
        "btn_nat"
    )

with col2:
    menu_column(
        "Mostra komunale",
        "Gjeneroni ndarjen e mostrës në nivel komune sipas vendbanimit dhe etnisë. Lejon aplikimin e filterëve demografikë dhe oversampling.",
        "pages/municipality-sample.py",
        "images/municipality.png",
        "btn_kom"
    )

with col3:
    menu_column(
        "Mostra për biznese",
        "Gjeneroni ndarjen e mostrës së bizneseve sipas komunës/regjionit dhe aktivitetit (NACE). Lejon aplikimin e filterëve të profilit të bizneseve dhe oversampling.",
        "pages/business-sample.py",
        "images/business.png",
        "btn_biz"
    )