import streamlit as st
import pandas as pd
import numpy as np
import base64
from io import BytesIO
from docx import Document

narrative_template_common = """
**Sampling Universe**

An up-to-date list of active businesses from the Kosovo Business Registration Agency (KBRA) will be used as the sampling universe. This database serves as the foundation for both the sample design and the population of the sample with businesses to be interviewed.

**Sampling Methodology**

The sampling methodology is a two-stage activity consisting of sample design and the selection of businesses to populate sample quotas. UBO Consulting is proposing a sample size of {N} businesses across Kosovo. The selection follows a rigorous approach to obtain a representative and unbiased survey sample, allowing findings to be generalized to the entire population of Kosovan businesses within a given power of statistical precision of a margin of error of {moe} at a 95% confidence interval.

**Stratification Strategy**
The survey employs a stratified random sampling strategy to ensure that the sample accurately mirrors the structure of the business population. This is achieved through a multi-stage stratification process:

- **Stage I: Geographic Stratification** - The population is first stratified into {first_strata} sub-groups. Sample sizes are determined based on the distribution of the business population, maintaining Probability Proportionate to Size (PPS).
"""

narrative_template_oversampling = """
**Oversampling Procedures** 

This survey employs a deliberate oversampling strategy to ensure robust data quality across all key analytical categories. For this, we apply **Disproportionate Stratified Sampling**:

- Base Quota: The general population is sampled according to its natural distribution.
- Boost Quota: The [Insert Target Group] is assigned a forced minimum quota (oversample) to ensure enough observations are collected to allow for valid comparisons between [Insert Comparison Categories, e.g., size classes / regions / sectors].
Adjustments for this oversampling will be handled during the analytical phase using post-stratification weights. This ensures that while the sample contains enough members of the target group for detailed analysis, the final results are weighted back to reflect the actual proportions of the Kosovo economy.
"""

narrative_template_randomization = """
**Sample Selection and Randomization**

In the final stage, the quotas for each basket type are populated using a **systematic random selection** technique. A randomization algorithm is employed within each basket to ensure that the selected firms mirror the structure of the populations as closely as possible.  
"""

narrative_template_reserves = """
**Replacement Strategy**

A reserve list of businesses for each basket is generated to serve as potential replacements during data collection. Recognizing that administrative lists may contain inactive businesses or inaccurate contact details—or that representatives may refuse participation—reserve lists will be used strictly as equivalent replacements (within the exact same basket/stratum) to maintain sample integrity. To ensure fieldwork continuity without the need for re-sampling, the reserve list is generated with a size of {reserve_size}, depending on the anticipated non-response and ineligibility rates for each specific stratum.
"""

# =====================================================
# HELPER FUNCTIONS
# =====================================================

def controlled_rounding(values: np.ndarray,
                        total_n: int,
                        seed: int = 42) -> np.ndarray:
    vals = np.asarray(values, dtype=float)
    if len(vals) == 0:
        return vals.astype(int)

    floors = np.floor(vals).astype(int)
    diff = int(total_n - floors.sum())

    if diff == 0:
        return floors

    fracs = vals - floors
    rng = np.random.default_rng(seed)

    if diff > 0:
        if fracs.sum() == 0:
            indices = rng.choice(len(vals), size=diff, replace=False)
        else:
            probs = fracs / fracs.sum()
            indices = rng.choice(len(vals), size=diff, replace=False, p=probs)
        floors[indices] += 1

    elif diff < 0:
        diff = -diff
        order = np.argsort(fracs)
        indices = order[:diff]
        floors[indices] -= 1

    final_diff = int(total_n - floors.sum())
    if final_diff != 0:
        idx = rng.integers(0, len(vals))
        floors[idx] += final_diff

    return floors


def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=True, sheet_name=sheet_name)
    return output.getvalue()


def create_download_link(file_bytes: bytes, filename: str, label: str):
    b64 = base64.b64encode(file_bytes).decode()
    button_html = f"""
    <a href="data:application/octet-stream;base64,{b64}" download="{filename}" style="text-decoration:none;">
        <div style="
            background-color:#344b77;
            color:white;
            text-align:center;
            font-weight:500;
            font-size:16px;
            padding:10px;
            border-radius:8px;
            margin-top:8px;
            width:100%;
            cursor:pointer;">
        {label}
        </div>
    </a>
    """
    st.markdown(button_html, unsafe_allow_html=True)


def create_download_link2(file_bytes: bytes, filename: str, label: str):
    b64 = base64.b64encode(file_bytes).decode()
    button_html = f"""
    <a href="data:application/octet-stream;base64,{b64}" download="{filename}" style="text-decoration:none;">
        <div style="
            background-color:#5b8fb8;
            color:white;
            text-align:center;
            font-weight:500;
            font-size:16px;
            padding:10px;
            border-radius:8px;
            margin-top:8px;
            width:100%;
            cursor:pointer;">
        {label}
        </div>
    </a>
    """
    st.markdown(button_html, unsafe_allow_html=True)


def narrative_to_word(text: str) -> bytes:    
    doc = Document()
    for line in text.split("\n"):
        doc.add_paragraph(line)
    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

def generate_recode_spss():
    return r"""
    RECODE D3 (2=7)
    (3=4)
    (8=1)
    (4=5)
    (5=1)
    (6=7)
    (7=6)
    (9=1)
    (10=5)
    (11=3)
    (12=7)
    (13=5)
    (14=6)
    (15=3)
    (16=6)
    (17=2)
    (18=1)
    (19=4)
    (20=4)
    (21=2)
    (22=2)
    (23=1)
    (24=1)
    (25=6)
    (26=3)
    (27=1)
    (1=1)
    (28=4)
    (29=7)
    (30=6)
    (33=5)
    (34=5)
    (31=2)
    (32=4)
    (35=6)
    (36=2)
    (37=2)
    (38=2)
    INTO Regjioni.
    VARIABLE LABELS  Regjioni 'Regjioni'.
    EXECUTE.

    * RECODE - Madhësia e biznesit (sipas numrit të punëtorëve).
    RECODE NumriPuntoreve 
        (0 THRU 9 = 1)
        (10 THRU 49 = 2)
        (50 THRU 249 = 3)
        (250 THRU HI = 4)
        INTO BizSize.
    VARIABLE LABELS BizSize 'Madhësia e biznesit sipas punëtorëve'.
    VALUE LABELS BizSize
        1 'Mikrondërmarrje'
        2 'Ndërmarrje e vogël'
        3 'Ndërmarrje e mesme'
        4 'Ndërmarrje e madhe'.
    EXECUTE.
    """

def generate_weighting_spss_syntax(wdf):
    """
    Merr tabelën wdf dhe prodhon sintaksën RAKE për SPSS.
    Pret që wdf të përmbajë kolonat:
    - Dimensioni
    - Kategoria
    - Kodi
    - Pesha
    """

    syntax = "* SPSS WEIGHTING SYNTAX FOR BUSINESS SURVEY.\n\n"

    # Shto recode për madhësinë e biznesit
    syntax += generate_recode_spss()

    syntax += "\n* RAKE WEIGHTING.\n"
    syntax += "SPSSINC RAKE\n"

    dim_index = 1

    for dim in wdf["Dimensioni"].unique():

        df_dim = wdf[wdf["Dimensioni"] == dim]

        if df_dim.empty:
            continue

        syntax += f"DIM{dim_index}={dim}\n"

        for _, row in df_dim.iterrows():
            code = int(row["Kodi"])
            coef = float(row["Pesha"])
            syntax += f"  {code} {coef}\n"

        dim_index += 1

    syntax += "FINALWEIGHT = peshat.\nEXECUTE.\n"

    return syntax

def add_codes_to_business_wdf(wdf):
    """
    Marrim wdf:
        Dimensioni | Kategoria | Populacioni | Pesha
    Dhe i shtojmë kolonën 'Kodi' sipas logjikës së peshimit të bizneseve.
    """

    # ============================
    # Kodet e Regjioneve
    # ============================
    region_codes = {
        "Regjioni i Prishtinës": 1,
        "Regjioni i Mitrovicës": 2,
        "Regjioni i Pejës": 3,
        "Regjioni i Prizrenit": 4,
        "Regjioni i Ferizajit": 5,
        "Regjioni i Gjilanit": 6,
        "Regjioni i Gjakovës": 7
    }

    # ============================
    # Kodet e Madhësisë së biznesit
    # ============================
    bizsize_codes = {
        "Mikrondërmarrje": 1,
        "Ndërmarrje e vogël": 2,
        "Ndërmarrje e mesme": 3,
        "Ndërmarrje e madhe": 4,
    }

    # ============================
    # Kodet e Sektorit (auto)
    # ============================
    # Krijon kodet dinamike 1..K sipas rendit alfabetik
    sectors = sorted(wdf[wdf["Dimensioni"] == "Sektori"]["Kategoria"].unique())
    sector_codes = {sec: i+1 for i, sec in enumerate(sectors)}

    # ============================
    # Inicializojmë kolonën Kodi
    # ============================
    wdf["Kodi"] = None

    # ============================
    # Loop për t’i caktuar kodet
    # ============================
    for i, row in wdf.iterrows():
        dim = row["Dimensioni"]
        cat = row["Kategoria"]

        if dim == "Regjioni":
            wdf.at[i, "Kodi"] = region_codes.get(cat)

        elif dim == "Madhësia e biznesit":
            wdf.at[i, "Kodi"] = bizsize_codes.get(cat)

        elif dim == "Sektori":
            wdf.at[i, "Kodi"] = sector_codes.get(cat)

    return wdf


# =====================================================
# CONFIG & PAGE STYLE
# =====================================================

st.set_page_config(
    page_title="Dizajnimi i Mostrës për Biznese",
    layout="wide"
)


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
            
.card {
    width: 100%;
    padding: 15px 20px;
    border-radius: 12px;
    background-color: #ffffff;
    border: 1px solid #e6e6e6;
    box-shadow: 0 1px 3px rgba(0,0,0,0.07);
    margin-bottom: 10px;
    }
.card-title {
    font-size: 18px;
    font-weight: 600;
    color: #344b77;
    margin-bottom: 8px;
    display: flex;
    align-items: center;
    gap: 8px;
    }
.card-title svg {
    width: 20px;
    height: 20px;
    }
.card-value {
    font-size: 16px;
    color: #000000;
    margin-bottom: 4px;
    }   
            
/* Set sidebar width */
[data-testid="stSidebar"] {
    width: 25% !important;
    min-width: 25% !important;
}

</style>
""", unsafe_allow_html=True)

st.title("Dizajnimi i Mostrës për Biznese")

# =====================================================
# LOAD BUSINESS DATA
# =====================================================

@st.cache_data(show_spinner="Duke ngarkuar të dhënat e bizneseve...")
def load_business_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)

    # Only active businesses
    if "Statusi" in df.columns:
        df = df[df["Statusi"].astype(str).str.lower() == "aktiv"]

    # Clean text columns
    for c in ["Komuna", "Regjioni", "Sektori", "NACE", "Forma juridike", "Madhësia e biznesit"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    return df

# Helper për alokim proporcional në një maskë
def alloc_to_mask(mask: pd.Series, quota: int) -> pd.Series:
    out = pd.Series(0.0, index=grouped.index)
    quota = int(quota)
    if quota <= 0:
        return out

    df_sub = grouped[mask].copy()
    if df_sub.empty:
        return out

    weights = df_sub["Pop_stratum"] / df_sub["Pop_stratum"].sum()
    floats = weights * quota
    ints = controlled_rounding(floats.to_numpy(), quota, seed)
    out.loc[df_sub.index] = ints
    return out

try:
    df_biz = load_business_data("excel-files/ARBK-bizneset.xlsx")
except Exception as e:
    st.error(f"Gabim gjatë leximit të regjistrit të bizneseve: {e}")
    st.stop()

if df_biz.empty:
    st.error("Regjistri i bizneseve është bosh.")
    st.stop()

# Remove unknown municipalities
df_biz = df_biz[df_biz["Komuna"].str.lower() != "i panjohur"]

# =====================================================
# SIDEBAR
# =====================================================

st.sidebar.header("Parametrat kryesorë")

n_total = st.sidebar.number_input(
    "Numri total i intervistave",
    min_value=1,
    value=500,
    step=100
)

first_strata = st.sidebar.selectbox(
    "Zgjidh variablën për ndarjen e parë",
    ["Regjioni", "Komuna"],
    index=1
)

second_strata = st.sidebar.multiselect(
    "Zgjidhni variablat për ndarjen e dytë",
    ["Madhësia e biznesit", "Sektori", "NACE", "Forma juridike"],
    default=["Madhësia e biznesit", "Sektori"]
)

strata_vars = [first_strata] + second_strata

survey_type = st.sidebar.selectbox(
    "Mbledhja e të dhënave",
    ["CAPI", "CATI", "CAWI"]
)

st.sidebar.markdown("---")

# =====================================================
# ADVANCED FILTERING SECTION
# =====================================================

st.sidebar.subheader("Filtrimi i bizneseve (opsionale)")

filter_options = [
    "Regjioni", 
    "Komuna", 
    "Madhësia e biznesit", 
    "NACE", 
    "Aktivitetet",
    "Forma juridike",
    "Sektori"
]

selected_filters = st.sidebar.multiselect(
    "Zgjidh fushat që dëshiron të filtrosh",
    filter_options
)

# BEJ NJE KOPJE QE TE MOS PREKIM ORIGJINALIN
df_filtered = df_biz.copy()

# -----------------------------
# FILTER: REGJIONI
# -----------------------------
if "Regjioni" in selected_filters:
    unique_vals = sorted(df_biz["Regjioni"].dropna().unique())
    selected_vals = st.sidebar.multiselect(
        "Zgjidh Regjionet",
        unique_vals
    )
    if selected_vals:
        df_filtered = df_filtered[df_filtered["Regjioni"].isin(selected_vals)]

# -----------------------------
# FILTER: KOMUNA
# -----------------------------
if "Komuna" in selected_filters:
    unique_vals = sorted(df_biz["Komuna"].dropna().unique())
    selected_vals = st.sidebar.multiselect(
        "Zgjidh Komunat",
        unique_vals
    )
    if selected_vals:
        df_filtered = df_filtered[df_filtered["Komuna"].isin(selected_vals)]

# -----------------------------
# FILTER: KATEGORIA
# -----------------------------
if "Madhësia e biznesit" in selected_filters:
    unique_vals = sorted(df_biz["Madhësia e biznesit"].dropna().unique())
    selected_vals = st.sidebar.multiselect(
        "Zgjidh Kategorinë",
        unique_vals
    )
    if selected_vals:
        df_filtered = df_filtered[df_filtered["Madhësia e biznesit"].isin(selected_vals)]

# -----------------------------
# FILTER: NACE
# -----------------------------
if "NACE" in selected_filters:
    unique_vals = sorted(df_biz["NACE"].dropna().unique())
    selected_vals = st.sidebar.multiselect(
        "Zgjidh NACE kodet",
        unique_vals
    )
    if selected_vals:
        df_filtered = df_filtered[df_filtered["NACE"].isin(selected_vals)]

# -----------------------------
# FILTER: AKTIVITETET
# -----------------------------
if "Aktivitetet" in selected_filters:
    unique_vals = sorted(df_biz["Aktivitetet"].dropna().unique())
    selected_vals = st.sidebar.multiselect(
        "Zgjidh Aktivitetet",
        unique_vals
    )
    if selected_vals:
        df_filtered = df_filtered[df_filtered["Aktivitetet"].isin(selected_vals)]

# -----------------------------
# FILTER: TIPI I BIZNESIT
# -----------------------------
if "Forma juridike" in selected_filters:
    unique_vals = sorted(df_biz["Forma juridike"].dropna().unique())
    selected_vals = st.sidebar.multiselect(
        "Zgjidh formën juridike",
        unique_vals
    )
    if selected_vals:
        df_filtered = df_filtered[df_filtered["Forma juridike"].isin(selected_vals)]

# -----------------------------
# FILTER: SEKTORI
# -----------------------------
if "Sektori" in selected_filters:
    unique_vals = sorted(df_biz["Sektori"].dropna().unique())
    selected_vals = st.sidebar.multiselect(
        "Zgjidh Sektorin",
        unique_vals
    )
    if selected_vals:
        df_filtered = df_filtered[df_filtered["Sektori"].isin(selected_vals)]


# =====================================================
# REPLACE MAIN POPULATION WITH FILTERED VERSION
# =====================================================
if df_filtered.empty:
    st.error("Nuk ka asnjë biznes pas filtrave të zgjedhur! Zgjidh filtrime më të gjera.")
    st.stop()

df_biz = df_filtered

st.sidebar.markdown("---")

st.sidebar.subheader("Rezervat për kontakte")

reserve_mode = st.sidebar.radio(
    "Metoda për llogaritjen e rezervave:",
    ["Përqindje (%)", "Proporcion"],
    index=0
)

reserve_percentage = None
reserve_ratio = None

if reserve_mode == "Përqindje (%)":
    reserve_percentage = st.sidebar.number_input(
        "Shkruaj përqindjen e rezervave (%)",
        min_value=1,
        max_value=500,
        value=20,
        step=10
    )

elif reserve_mode == "Proporcion":
    reserve_ratio = st.sidebar.number_input(
        "Vendos numrin për proporcion (p.sh. 2 për 2:1)",
        min_value=1,
        max_value=10,
        value=2,
        step=1
    )


st.sidebar.markdown("---")

oversample_enabled = st.sidebar.checkbox("Oversampling", value=False)
oversample_inputs = {}   # {var: [{value, n}, ...], ...}

if oversample_enabled:

    oversample_vars = st.sidebar.multiselect(
        "Zgjidh deri në 2 variabla për oversample:",
        ["Komuna", "Madhësia e biznesit", "Sektori"],
        max_selections=2
    )

    for var in oversample_vars:

        st.sidebar.markdown(f"### {var}")

        # get valid categories for that variable
        valid_vals = df_biz[var].dropna().unique().tolist()

        selected_vals = st.sidebar.multiselect(
            f"Zgjidh vlerat për {var}",
            valid_vals,
            key=f"multi_{var}"
        )

        entry_list = []
        for v in selected_vals:
            q = st.sidebar.number_input(
                f"Kuota për {var} = {v}",
                min_value=1, value=50, key=f"quota_{var}_{v}"
            )
            entry_list.append({"value": v, "n": q})

        oversample_inputs[var] = entry_list
else:
    oversample_vars = []

run_button = st.sidebar.button("Gjenero shpërndarjen e mostrës")

# =====================================================
# MAIN LOGIC
# =====================================================

if run_button:

    caption_main = (
        f"Totali i mostrës: **{n_total}**"
    )

    # 1) BUILD STRATA POPULATION
    grouped = (
        df_biz.groupby(strata_vars)
        .size()
        .reset_index(name="Pop_stratum")
    )

    total_pop = grouped["Pop_stratum"].sum()

    # =====================================================
    # CALCULATE MARGIN OF ERROR (95% confidence)
    # =====================================================

    z = 1.96        # z-score for 95% CI
    p = 0.5         # worst-case proportion
    n = n_total     # desired sample size
    Npop = total_pop   # population size after filters

    # Finite Population Correction
    if Npop > n:
        fpc = np.sqrt((Npop - n) / (Npop - 1))
    else:
        fpc = 1.0

    moe = z * np.sqrt((p * (1 - p)) / n) * fpc
    moe_percent = moe * 100


    grouped["n_alloc"] = 0
    seed = 42

    # ============================
    # OVERSAMPLING LOGIC (supports 0,1,2 variables)
    # ============================
    grouped["n_alloc"] = 0
    seed = 42

    # Build full OS list from oversample_inputs
    all_os = []

    for var, entry_list in oversample_inputs.items():
        for entry in entry_list:
            mask = (grouped[var] == entry["value"])
            all_os.append({
                "var": var,
                "value": entry["value"],
                "n": entry["n"],
                "mask": mask
            })

    # ----------------------------------------
    # CASE 0 — No oversampling
    # ----------------------------------------
    if len(all_os) == 0:
        weights = grouped["Pop_stratum"] / grouped["Pop_stratum"].sum()
        floats = weights * n_total
        grouped["n_alloc"] = controlled_rounding(floats.to_numpy(), n_total, seed)

    # ----------------------------------------
    # CASE 1 — Single oversample variable/value
    # ----------------------------------------
    elif len(all_os) == 1:

        osA = all_os[0]

        # allocate OS group
        alloc_A = alloc_to_mask(osA["mask"], osA["n"])
        grouped["n_alloc"] += alloc_A

        # allocate rest proportionally outside OS mask
        remaining = n_total - int(alloc_A.sum())
        mask_rest = (grouped["n_alloc"] == 0)
        alloc_rest = alloc_to_mask(mask_rest, remaining)
        grouped["n_alloc"] += alloc_rest

    # ----------------------------------------
    # CASE 2 — Two oversample variables
    # ----------------------------------------
    else:
        # sort by quota descending
        all_os_sorted = sorted(all_os, key=lambda x: x["n"], reverse=True)

        # largest quota = OS-B
        osB = all_os_sorted[0]
        osA_list = all_os_sorted[1:]   # the remaining OS groups

        # Step 1 – allocate B completely
        alloc_B = alloc_to_mask(osB["mask"], osB["n"])
        grouped["n_alloc"] = alloc_B

        # Step 2 – allocate A variables one by one
        for osA in osA_list:

            # overlap between A and B
            overlap_mask = osA["mask"] & osB["mask"]
            overlap_from_B = int(alloc_B[overlap_mask].sum())

            # remaining A quota after removing overlap count
            remaining_A = max(osA["n"] - overlap_from_B, 0)

            # allocate A only where B has NOT allocated
            alloc_A = alloc_to_mask(osA["mask"] & ~osB["mask"], remaining_A)
            grouped["n_alloc"] += alloc_A

        # Step 3 – allocate the rest outside any OS mask
        used = int(grouped["n_alloc"].sum())
        remaining = max(n_total - used, 0)

        combined_mask = sum([os["mask"] for os in all_os]) > 0
        mask_rest = ~combined_mask

        alloc_rest = alloc_to_mask(mask_rest, remaining)
        grouped["n_alloc"] += alloc_rest

    total_alloc = grouped["n_alloc"].sum()

    if selected_filters:
        filters_text = ", ".join(selected_filters)
    else:
        filters_text = "Asnjë"

    strata_text = ", ".join(strata_vars)

    if oversample_enabled:
        oversampling_text = ", ".join(oversample_inputs)
    else:
        oversampling_text = "Joaktiv"

    def load_svg(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()

    icon_sample = load_svg("images/sample-business.svg")
    icon_strata = load_svg("images/strata.svg")

    col1, col2 = st.columns(2)

    with col1:
        with st.container():
            st.markdown(f"""
            <div class='card'>
                <div class='card-title'>
                    {icon_sample} Mostra
                </div>
                <div class='card-value'>Totali i mostrës: <b>{n_total}</b></div>
                <div class='card-value'>Marzha e gabimit: <b>± {moe_percent:.2f}%</b></div>
                <div class='card-value'>Intervali i besimit: <b>95%</b></div>
            </div>
            """, unsafe_allow_html=True)

    with col2:
        with st.container():
            st.markdown(f"""
            <div class='card'>
                <div class='card-title'>
                        {icon_strata} Ndarja e mostrës
                </div>
                <div class='card-value'>Ndarja sipas: <b>{strata_text}</b></div>
                <div class='card-value'>Filtrimi sipas: <b>{filters_text}</b></div>
                <div class='card-value'>Oversampling: <b>{oversampling_text}</b></div>
            </div>
            """, unsafe_allow_html=True)

    st.subheader("Tabela e ndarjes së mostrës")
    df_final = grouped[grouped["n_alloc"] > 0].copy()
    df_final = df_final.drop(columns=["Pop_stratum"])
    df_final = df_final.rename(columns={"n_alloc": "Intervista"})
    
    total_row = pd.DataFrame([{
        **{col: "Total" for col in strata_vars},
        "Intervista": df_final["Intervista"].sum()
    }])

    df_final = pd.concat([df_final, total_row], ignore_index=True)

    st.caption(caption_main)
    st.dataframe(df_final, use_container_width=True)

    excel_bytes_reserve = df_to_excel_bytes(
            df_final, 
            sheet_name="Rezervat_Biznese"
        )

    create_download_link(
            file_bytes=excel_bytes_reserve,
            filename="mostra_biznese.xlsx",
            label="Shkarko Mostrën"
        )

    # =====================================================
    # REZERVAT – LISTË BIZNESESH 2× PER STRATUM
    # =====================================================

    st.markdown("---")
    st.subheader("Lista e kontakteve të bizneseve")

    reserve_rows = []
    warnings_list = []  # për të shfaqur info për mungesë të kontakteve

    for idx, row in df_final[:-1].iterrows():  # EXCLUDE total row
        intervista = row["Intervista"]
        if pd.isna(intervista) or intervista == "":
            continue

        intervista = int(intervista)
        # ----------------------------------------------------
        # RESERVE ALLOCATION BASED ON SELECTED MODE
        # ----------------------------------------------------
        if  reserve_mode == "Përqindje (%)":
            reserve_n = int(intervista * (reserve_percentage / 100))

        elif reserve_mode == "Proporcion":
            reserve_n = int(intervista * (reserve_ratio - 1))

        # Minimum 1 if reserve exists but evaluates 0
        if reserve_n < 0:
            reserve_n = 0
            st.warning("Zgjedhni një vlerë të saktë për rezerva")

        # ----------------------------------------------------
        # Filter original DF based on strata variables
        # ----------------------------------------------------
        mask = pd.Series(True, index=df_biz.index)
        for col in strata_vars:
            mask &= (df_biz[col] == row[col])

        df_stratum = df_biz[mask].copy()


        # Filter only businesses with valid phone number
        df_stratum = df_stratum[df_stratum["Numri i telefonit"].notnull() & (df_stratum["Numri i telefonit"] != "")]

        if df_stratum.empty:
            warnings_list.append(
                f"Nuk ka asnjë biznes me numër telefoni të vlefshëm për stratum-in: {row.to_dict()}"
            )
            continue

        available_n = len(df_stratum)

        # ----------------------------------------------------
        # Sampling for reserve list
        # ----------------------------------------------------
        if available_n < reserve_n:
            warnings_list.append(
                f"Mungesë kontakti: për stratum-in {row.to_dict()} "
                f"duhen {reserve_n} biznese, por janë vetëm {available_n} me numër telefoni."
            )
            reserve_sample = df_stratum.sample(
                n=available_n,
                replace=False,
                random_state=42
            )
        else:
            reserve_sample = df_stratum.sample(
                n=reserve_n,
                replace=False,
                random_state=42
            )

        # Add strata label for debugging/clarity
        reserve_sample["Strata"] = " | ".join([
            f"{col}: {row[col]}" for col in strata_vars
        ])

        reserve_sample["Rezervë_kërkuar"] = reserve_n
        reserve_sample["Rezervë_marrë"] = len(reserve_sample)

        reserve_rows.append(reserve_sample)

    # =====================================================
    # OUTPUT – Tabela e rezervave
    # =====================================================

    if reserve_rows:
        df_reserves = pd.concat(reserve_rows, ignore_index=True)

        # Output columns you requested
        reserve_cols = [
            "Emri i biznesit",
            "Numri i telefonit",
            "Komuna",
            "Regjioni",
            "Forma juridike",
            "Madhësia e biznesit",
            "Sektori",
            "Aktivitetet"
        ]

        reserve_cols = [c for c in reserve_cols if c in df_reserves.columns]

        with st.expander("Shfaq tabelën e plotë të kontakteve", expanded=False):     
            st.dataframe(df_reserves[reserve_cols], use_container_width=True)

        # DOWNLOAD
        excel_bytes_reserve = df_to_excel_bytes(
            df_reserves[reserve_cols], 
            sheet_name="Kontaktet_Biznese"
        )

        create_download_link(
            file_bytes=excel_bytes_reserve,
            filename="lista_kontakteve_biznese.xlsx",
            label="Shkarko Listën e Kontakteve"
        )

    else:
        st.info("Nuk u gjeneruan rezerva për asnjë stratum.")

    # =====================================================
    # SHFAQ MUNGESAT E KONTAKTEVE
    # =====================================================

    if warnings_list:
        st.warning("Disa strata nuk kanë mjaftueshëm biznese me numra telefoni.")


    # =====================================================
    # WEIGHTING TABLE
    # =====================================================
    weighting_vars = ["Regjioni", "Sektori", "Madhësia e biznesit"]

    # Remove oversample variables from weighting vars
    weighting_vars = [v for v in weighting_vars if v not in oversample_vars]

    if weighting_vars:
        df_biz = df_biz[df_biz["Regjioni"] != "I panjohur"]
        st.markdown("---")
        st.subheader("Sintaksa për peshim në SPSS")

        w_rows = []
        for var in weighting_vars:
            if var not in df_biz.columns:
                continue
            vc = df_biz[var].value_counts()
            for cat, pop in vc.items():
                w_rows.append({
                    "Dimensioni": var,
                    "Kategoria": cat,
                    "Populacioni": int(pop),
                    "Pesha": pop / vc.sum()
                })
        wdf = pd.DataFrame(w_rows)
        wdf = add_codes_to_business_wdf(wdf)
        
        with st.expander("Shfaq tabelën e plotë të peshave", expanded=False):
            st.dataframe(wdf, use_container_width=True)

        spss_text = generate_weighting_spss_syntax(wdf)

        create_download_link(
            file_bytes=spss_text.encode("utf-8-sig"),
            filename="peshat_biznese.sps",
            label="Shkarko Peshat për SPSS"
        )

    # =====================================================
    # NARRATIVE
    # =====================================================

    st.markdown("---")
    st.subheader("Përshkrimi i dizajnimit të mostrës")

    strata_text = ", ".join(strata_vars)

    narrative_text = f"""
Sample Design – Business Survey

Country: Kosovo  
Survey type: {survey_type}  
Target completes: {n_total}

Sampling Population  
All registered and **active** businesses with valid contact phone numbers.

Sampling Frame  
Official business register (ARBK), containing firm ID, region, municipality,  
legal entity type, NACE sector, size band, and contact information.

Stratification  
Strata used: **{strata_text}**.  
Interview allocation is proportional to the number of businesses in each stratum.

Oversampling  
{"Oversampling applied." if oversample_enabled else "No oversampling used."}
"""

    with st.expander("Shfaq narrativën"):
        st.markdown(narrative_text)

    narrative_doc = narrative_to_word(narrative_text)
    create_download_link(
        narrative_doc,
        "Narrativa_Mostra_Biznese.docx",
        "Shkarko Narrativën (Word)"
    )

else:
    st.info("Cakto parametrat kryesorë dhe kliko **'Gjenero shpërndarjen e mostrës'** për të dizajnuar mostrën.")
