import streamlit as st
import pandas as pd
import numpy as np
import pydeck as pdk

from pages.national_sample import (compute_filtered_pop_for_psu_row, controlled_rounding, load_psu_data, df_to_excel_bytes, 
                                   create_download_link, create_download_link2, compute_population_coefficients, add_codes_to_coef_df, df_eth, df_ga, region_map)

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

def generate_spss_syntax_municipality(coef_df, data_collection_method):

    out = "* Encoding: UTF-8.\n\n"

    # RECODE për Grupmosha
    if data_collection_method == "CAWI":
        out += (
            "RECODE D2 (MISSING=COPY) "
            "(18 THRU 24 = 1) "
            "(25 THRU 34 = 2) "
            "(35 THRU 44 = 3) "
            "(45 THRU 54 = 4) "
            "(55 THRU HI = 5) "
            "INTO Grupmoshat.\n"
        )
    else:
        out += (
            "RECODE D2 (MISSING=COPY) "
            "(18 THRU 24 = 1) "
            "(25 THRU 34 = 2) "
            "(35 THRU 44 = 3) "
            "(45 THRU 54 = 4) "
            "(55 THRU 64 = 5) "
            "(65 THRU HI = 6) "
            "INTO Grupmoshat.\n"
        )

    out += "\nSPSSINC RAKE\n"

    dim_order = ["Gjinia", "Grupmosha", "Vendbanimi", "Etnia"]

    i = 1
    for dim in dim_order:
        df_dim = coef_df[coef_df["Dimensioni"] == dim]

        out += f"DIM{i}={dim} "
        for _, row in df_dim.iterrows():
            out += f"{int(row['Kodi'])} {row['Pesha']}\n"
        i += 1

    out += "FINALWEIGHT=peshat.\n"

    return out

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

data_collection_method = st.sidebar.selectbox(
    "Metoda e mbledhjes së të dhënave",
    options=["CAPI", "CATI", "CAWI"],
    index=0,
    key="Metoda për komuna"
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

    st.subheader("Tabela e ndarjes së mostrës brenda komunës")
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

    # =====================================================
    # 6) INTERACTIVE MAP WITH FOLIUM (Urban removed)
    # =====================================================

    st.subheader("Harta e vendeve të përzgjedhura në mostër")

        # Remove the artificial urban row BEFORE merging with coordinates
    df_map = final[["Komuna", "Fshati/Qyteti", "Intervista"]].copy()
    df_map.loc[df_map["Fshati/Qyteti"] == "Urban", "Fshati/Qyteti"] = \
    df_map.loc[df_map["Fshati/Qyteti"] == "Urban", "Komuna"]


        # Merge with PSU coordinates
    df_map = df_map.merge(
            df_psu[["Komuna", "Fshati/Qyteti", "lat", "long"]],
            on=["Komuna", "Fshati/Qyteti"],
            how="left"
        )

    layer = pdk.Layer(
        "ScatterplotLayer",
        data=df_map,
        get_position='[long, lat]',
        get_fill_color='[200, 30, 0, 160]',
        get_radius=200,
        pickable=True
    )

    view_state = pdk.ViewState(
        latitude=df_map["lat"].mean(),
        longitude=df_map["long"].mean(),
        zoom=11
    )

    deck = pdk.Deck(
        layers=[layer],
        initial_view_state=view_state,
        map_provider="carto",     # ⭐ REQUIRED
        map_style="light",        # ⭐ WORKS WITHOUT TOKEN
        tooltip={"html": "<b>{Fshati/Qyteti}</b><br>{Intervista} intervista"}
    )

    st.pydeck_chart(deck)
    
    coef_df = compute_population_coefficients(
    df_ga=df_ga,
    df_eth=df_eth,
    region_map=region_map,
    gender_selected=gender_selected,
    min_age=min_age,
    max_age=max_age,
    eth_filter=eth_filter,
    settlement_filter=["Urban","Rural"],  # brenda komunes
    komuna_filter=[komuna],               # shumë e rëndësishme!
    data_collection_method=add_codes_to_coef_df
)

    dims_to_keep = ["Gjinia", "Grupmosha", "Vendbanimi", "Etnia"]
    coef_df = coef_df[coef_df["Dimensioni"].isin(dims_to_keep)]

    # Remove dimensions with only one category
    valid_dims = (
        coef_df.groupby("Dimensioni")["Kategoria"]
        .nunique()
    )

    dims_valid = valid_dims[valid_dims > 1].index.tolist()

    coef_df = coef_df[coef_df["Dimensioni"].isin(dims_valid)]

    coef_df = add_codes_to_coef_df(
    coef_df,
    data_collection_method
)

    spss_text = generate_spss_syntax_municipality(
    coef_df,
    data_collection_method
)
    st.markdown("---")
    st.subheader("Sintaksa për peshim në SPSS")

    with st.expander("Shfaq tabelën e plotë të peshave", expanded=False):
        st.dataframe(coef_df, use_container_width=True)

    create_download_link(
        file_bytes=spss_text.encode("utf-8"),
        filename=f"sintaksa_peshat_{komuna}.sps",
        label="Shkarko Peshat për SPSS"
    )




else:
    st.info("Cakto parametrat kryesorë dhe kliko **'Gjenero shpërndarjen e mostrës'** për të dizajnuar mostrën.")
