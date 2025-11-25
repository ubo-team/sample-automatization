import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import base64
import re
from pages.national_sample import (compute_filtered_pop_for_psu_row, controlled_rounding, load_psu_data, df_to_excel_bytes, create_download_link, create_download_link2)

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

/* Hide only page navigation links, keep widgets **/
[data-testid="stSidebarNav"] li {
    display: none !important;
}

/* Keep sidebar header + widgets */
[data-testid="stSidebarNav"] {
    padding-bottom: 0 !important;
}

</style>
""", unsafe_allow_html=True)

# =====================================================
# PAGE SETTINGS & HEADER
# =====================================================

st.set_page_config(
    page_title="Mostra sipas Komunës",
    layout="wide"
)

# =====================================================
# LOAD PSU DATA
# =====================================================

df_psu = load_psu_data("excel-files/ASK-2024-Komuna-Vendbanim-Fshat+Qytet.xlsx")  # :contentReference[oaicite:1]{index=1}

municipalities = sorted(df_psu["Komuna"].unique())

# =====================================================
# SIDEBAR UI
# =====================================================

st.sidebar.header("Parametrat kryesorë")

komuna = st.sidebar.selectbox(
    "Zgjidh Komunën",
    options=municipalities,
    index=municipalities.index("Prishtinë") if "Prishtinë" in municipalities else 0
)

N = st.sidebar.number_input(
    "Numri i mostrës për komunen",
    min_value=6,
    value=800,
    step=2
)

st.sidebar.markdown("---")
st.sidebar.subheader("Filtrat demografikë")

gender_selected = st.sidebar.multiselect(
    "Gjinia",
    ["Meshkuj", "Femra"],
    default=["Meshkuj", "Femra"]
)

min_age = st.sidebar.number_input("Mosha minimale", min_value=0, value=18)

max_age_input = st.sidebar.text_input("Mosha maksimale (opsionale)")
max_age = int(max_age_input) if max_age_input.strip() else None

eth_filter = st.sidebar.multiselect(
    "Etnitë që përfshihen",
    options=["Shqiptar", "Serb", "Tjerë"],
    default=["Shqiptar", "Serb", "Tjerë"], 
    key = "Etnia"
)

st.sidebar.markdown("---")

run = st.sidebar.button("Gjenero shpërndarjen e mostrës", key="generate_sample_button")

# =====================================================
# MAIN LOGIC
# =====================================================

if run:
    st.title(f"Mostra e Komunës - {komuna}")

    # Subset for municipality
    df_mun = df_psu[df_psu["Komuna"] == komuna].copy()

    eth_other = ["Boshnjak", "Turk", "Rom", "Ashkali", "Egjiptian", "Goran", "Të tjerë"]
    
    if eth_filter == "Tjerë":
        eth_filter.remove("Tjerë")
        eth_filter.extend(eth_other)

    # Compute filtered population (PopFilt)
    df_mun["PopFilt"] = df_mun.apply(
        lambda r: compute_filtered_pop_for_psu_row(
            psu_row=r,
            age_min=min_age,
            age_max=max_age,
            gender_selected=gender_selected,
            eth_filter=eth_filter
        ),
        axis=1
    )

    # Split Urban / Rural
    df_urban = df_mun[df_mun["Vendbanimi"] == "Urban"].copy()
    df_rural = df_mun[df_mun["Vendbanimi"] == "Rural"].copy()

    urban_pop = df_urban["PopFilt"].sum()
    rural_pop = df_rural["PopFilt"].sum()

    total_pop = urban_pop + rural_pop

    if total_pop == 0:
        st.error("Popullsia pas filtrimit është zero. Ndrysho filtrat.")
        st.stop()

    # =====================================================
    # 1) Allocate N into Urban & Rural proportionally
    # =====================================================

    floats_ur = np.array([
        N * (urban_pop / total_pop),
        N * (rural_pop / total_pop)
    ])

    urban_n, rural_n = controlled_rounding(floats_ur, N)

    # =====================================================
    # 2) Urban: always 1 row
    # =====================================================

    if not df_urban.empty:
        df_urban = df_urban.iloc[[0]].copy()   # always 1 urban
        df_urban["Intervista"] = urban_n
        df_urban["Fshati/Qyteti"] = "Urban"
    else:
        df_urban["Intervista"] = 0

    # =====================================================
    # 3) Rural distribution with loop until last ≥ 6
    # =====================================================

    def distribute_rural(df_rural, rural_n):
        df_rural = df_rural.copy()

        while True:
            weights = df_rural["PopFilt"] / df_rural["PopFilt"].sum()
            floats = weights * rural_n
            alloc = controlled_rounding(floats, rural_n)

            df_rural["Intervista"] = alloc
            df_rural = df_rural.sort_values("Intervista", ascending=False)

            last_value = df_rural["Intervista"].iloc[-1]

            if last_value >= 6:
                return df_rural

            # Remove last village and try again
            df_rural = df_rural.iloc[:-1]

            if df_rural.empty:
                st.error("Asnjë fshat nuk mund të marrë minimum 6 intervista.")
                st.stop()

    if rural_n > 0 and not df_rural.empty:
        df_rural_final = distribute_rural(df_rural, rural_n)
    else:
        df_rural_final = df_rural.copy()
        df_rural_final["Intervista"] = 0

    # =====================================================
    # 4) Combine Final Output
    # =====================================================

    final = pd.concat([df_urban, df_rural_final], ignore_index=True)

    sample = final[
        ["Fshati/Qyteti","Intervista"]
    ]

    sample.loc["Total"] = [
    "Total",                   # value for the first (string) column
    sample["Intervista"].sum() # sum of the numeric column
]

    strata = final[
        ["Komuna", "Vendbanimi", "Fshati/Qyteti", "PopFilt", "Intervista"]
    ]

    global_total = int(sample.loc["Total", "Intervista"])
    # Përgatit tekstin për grupmoshën
    if max_age is None:
        age_text = f"{min_age}+"
    else:
        age_text = f"{min_age}–{max_age}"

    caption_main = (
        f"Totali i mostrës: **{N}** | "
        f"Totali i alokuar: **{global_total}** | "
        f"Grupmosha: **{age_text}**"
    )

    st.caption(caption_main)

    st.subheader("Tabela finale e mostrës brenda komunës")
    st.dataframe(sample, use_container_width=True)

    with st.expander("Shfaq tabelën e plotë të stratum-eve (long format)", expanded=False):
        st.dataframe(strata, use_container_width=True)

    # =====================================================
    # 5) Download as Excel
    # =====================================================
    # 📘 Pivot table (Excel)
    pivot_excel = df_to_excel_bytes(sample, sheet_name="Mostra")
    create_download_link(
        file_bytes=pivot_excel,
        filename=f"mostra_e_gjeneruar_{komuna}.xlsx",
        label="Shkarko Mostrën"
    )

    # 📘 Strata table (Excel)
    strata_excel = df_to_excel_bytes(strata, sheet_name="Strata")
    create_download_link2(
        file_bytes=strata_excel,
        filename=f"mostra_strata_{komuna}.xlsx",
        label="Shkarko Strata"
    )
else:
    st.info("Cakto parametrat kryesorë dhe kliko **'Gjenero shpërndarjen e mostrës'** për të dizajnuar mostrën..")
