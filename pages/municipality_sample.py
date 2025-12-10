import streamlit as st
import base64
import pandas as pd
import numpy as np
import pydeck as pdk

# =====================================================
# PAGE SETTINGS & HEADER
# =====================================================

st.set_page_config(
    page_title="Mostra sipas Komunës",
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

@st.cache_data
def load_ethnicity_settlement_data(path: str) -> pd.DataFrame:
    """
    Load ASK-2024-Komuna-Etnia-Vendbanimi.xlsx
    Expected structure:
    Komuna | Vendbanimi | Shqiptar | Tjerë | Serb | ...
    Convert to long format: one row per (Komuna, Vendbanimi, Etnia).
    """
    df = pd.read_excel(path, sheet_name=0)
    # Identify ethnicity columns (all non-id cols except Komuna, Vendbanimi)
    id_cols = ["Komuna", "Vendbanimi"]
    eth_cols = [c for c in df.columns if c not in id_cols]
    df_long = df.melt(
        id_vars=id_cols,
        value_vars=eth_cols,
        var_name="Etnia",
        value_name="Pop_base"
    )
    # Clean
    df_long["Etnia"] = df_long["Etnia"].str.strip()
    df_long["Vendbanimi"] = df_long["Vendbanimi"].str.strip()
    df_long["Komuna"] = df_long["Komuna"].str.strip()
    df_long["Pop_base"] = df_long["Pop_base"].fillna(0).astype(float)
    return df_long


@st.cache_data
def load_gender_age_data(path: str) -> pd.DataFrame:
    """
    Load ASK-2024-Komuna-Gjinia-Mosha.xlsx (sheet 'census_00').
    Expected structure:
    Komuna | Gjinia | 0 | 1 | 2 | ... | (age columns)
    """
    df = pd.read_excel(path, sheet_name="census_00")
    # Normalize names
    df["Komuna"] = df["Komuna"].astype(str).str.strip()
    df["Gjinia"] = df["Gjinia"].astype(str).str.strip()

    # Keep only age columns that are pure ints as strings, e.g. "0","1",...
    age_cols = []
    for c in df.columns:
        s = str(c).strip()
        if s.isdigit():
            age_cols.append(c)

    # Remove empty rows (if any)
    df = df.dropna(subset=["Komuna", "Gjinia"])
    return df, age_cols


def get_region_mapping() -> dict:
    """
    Map Komuna -> Regjion (bazuar në ASK)
    """
    region_map = {
        "Deçan": "Gjakovë",
        "Dragash": "Prizren",
        "Ferizaj": "Ferizaj",
        "Fushë Kosovë": "Prishtinë",
        "Gjakovë": "Gjakovë",
        "Gjilan": "Gjilan",
        "Gllogoc": "Prishtinë",
        "Graçanicë": "Prishtinë",
        "Han i Elezit": "Ferizaj",
        "Istog": "Pejë",
        "Junik": "Gjakovë",
        "Kaçanik": "Ferizaj",
        "Kamenicë": "Gjilan",
        "Klinë": "Pejë",
        "Kllokot": "Gjilan",
        "Leposavic": "Mitrovicë",
        "Lipjan": "Prishtinë",
        "Malishevë": "Prizren",
        "Mamushë": "Prizren",
        "Mitrovicë": "Mitrovicë",
        "Mitrovica Veriore": "Mitrovicë",
        "Novobërdë": "Gjilan",
        "Obiliq": "Prishtinë",
        "Partesh": "Gjilan",
        "Pejë": "Pejë",
        "Podujevë": "Prishtinë",
        "Prishtinë": "Prishtinë",
        "Prizren": "Prizren",
        "Rahovec": "Gjakovë",
        "Ranillug": "Gjilan",
        "Shtërpcë": "Ferizaj",
        "Shtime": "Ferizaj",
        "Skënderaj": "Mitrovicë",
        "Suharekë": "Prizren",
        "Viti": "Gjilan",
        "Vushtrri": "Mitrovicë",
        "Zubin Potok": "Mitrovicë",
        "Zvecan": "Mitrovicë"
    }
    return region_map

def controlled_rounding(values: np.ndarray,
                        total_n: int,
                        seed: int = 42) -> np.ndarray:
    """
    Controlled rounding:
    - Start from float allocations
    - Floor all
    - Distribute remaining units based on fractional parts (probabilistic)
    - Sum-preserving & reproducible via seed
    """
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
        # Distribute +1 to 'diff' positions, weighted by fractional parts
        if fracs.sum() == 0:
            # if all fracs = 0, choose random indices uniformly
            indices = rng.choice(len(vals), size=diff, replace=False)
        else:
            probs = fracs / fracs.sum()
            indices = rng.choice(len(vals), size=diff, replace=False, p=probs)
        floors[indices] += 1

    elif diff < 0:
        # Too many after floor (rare, due to numeric issues).
        # Remove -diff units from smallest fractional parts.
        diff = -diff
        # Positions with smallest fracs get -1
        order = np.argsort(fracs)  # ascending
        indices = order[:diff]
        floors[indices] -= 1

    # Safety adjust if still off by 1 from numeric edge cases
    final_diff = int(total_n - floors.sum())
    if final_diff != 0 and len(vals) > 0:
        idx = rng.integers(0, len(vals))
        floors[idx] += final_diff

    return floors

@st.cache_data
def load_psu_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    # Normalizim minimal
    df["Komuna"] = df["Komuna"].astype(str).str.strip()
    df["Vendbanimi"] = df["Vendbanimi"].astype(str).str.strip()
    df["Fshati/Qyteti"] = df["Fshati/Qyteti"].astype(str).str.strip()
    df["Quadrant"] = df["Quadrant"].astype(str).str.strip()

    # Etnitë kryesore
    for col in ["Shqiptar", "Serb"]:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = df[col].fillna(0).astype(float)

    other_cols = [
        "Boshnjak", "Turk", "Rom", "Ashkali", "Egjiptian",
        "Goran", "Të tjerë", "Preferoj të mos përgjigjem"
    ]
    for col in other_cols:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = df[col].fillna(0).astype(float)

    df["Tjeter_pop"] = df[other_cols].sum(axis=1)

    return df

def compute_filtered_pop_for_psu_row(
    psu_row: pd.Series,
    age_min: int,
    age_max: int | None,
    gender_selected: list[str],
    eth_filter: list[str]
) -> float:
    """
    Compute population for one PSU using demographic filters.
    Fully safe: no division by zero, respects gender/age/eth filters.
    """

    import re

    # -------------------------------------------------------
    # 1. Handle age max
    # -------------------------------------------------------
    if age_max is None:
        age_max = 120  # large value = include all

    # -------------------------------------------------------
    # 2. Identify all PSU age group columns
    # -------------------------------------------------------
    age_cols = []
    for col in psu_row.index:
        name = str(col).strip()
        # match formats like '0-4', '5-9', '65+', etc.
        if re.fullmatch(r"\d+\-\d+", name) or re.fullmatch(r"\d+\+", name):
            age_cols.append(col)

    # Helper to decode age range
    def group_range(col_name):
        if "+" in col_name:
            base = int(col_name.replace("+", ""))
            return (base, 200)
        lo, hi = col_name.split("-")
        return (int(lo), int(hi))

    # -------------------------------------------------------
    # 3. AGE FILTER (fractional overlap)
    # -------------------------------------------------------
    pop_age = 0
    for col in age_cols:
        lo, hi = group_range(col)
        group_pop = psu_row[col]

        if group_pop <= 0:
            continue

        # overlap between age group and filter
        overlap_lo = max(lo, age_min)
        overlap_hi = min(hi, age_max)

        if overlap_lo > overlap_hi:
            continue  # no overlap

        group_size = hi - lo + 1
        overlap_size = overlap_hi - overlap_lo + 1

        fraction = overlap_size / group_size

        pop_age += group_pop * fraction

    # If no ages match → return 0
    if pop_age <= 0:
        return 0

    # -------------------------------------------------------
    # 4. GENDER FILTER (SAFE)
    # -------------------------------------------------------
    male = psu_row.get("Meshkuj", 0)
    female = psu_row.get("Femra", 0)

    if "Meshkuj" not in gender_selected:
        male = 0
    if "Femra" not in gender_selected:
        female = 0

    total_gender_pop = male + female

    if total_gender_pop <= 0:
        return 0

    # Proportion from allowed genders
    gender_fraction = total_gender_pop / (psu_row.get("Meshkuj", 0) + psu_row.get("Femra", 0)
                                          if (psu_row.get("Meshkuj", 0) + psu_row.get("Femra", 0)) > 0
                                          else total_gender_pop)

    # -------------------------------------------------------
    # 5. ETHNICITY FILTER (SAFE)
    # -------------------------------------------------------
    eth_total = 0
    eth_selected = 0

    all_ethnic_cols = [
        "Shqiptar", "Serb", "Boshnjak", "Turk", "Rom",
        "Ashkali", "Egjiptian", "Goran", "Të tjerë",
        "Preferoj të mos përgjigjem"
    ]

    for eth in all_ethnic_cols:
        eth_total += psu_row.get(eth, 0)

    for eth in eth_filter:
        eth_selected += psu_row.get(eth, 0)

    # SAFE division
    if eth_total > 0:
        eth_fraction = eth_selected / eth_total
    else:
        eth_fraction = 0

    if eth_fraction <= 0:
        return 0

    # -------------------------------------------------------
    # 6. Combine all three filters (age × gender × ethnicity)
    # -------------------------------------------------------
    final_pop = pop_age * (total_gender_pop / (male + female) if (male + female) > 0 else 1)
    final_pop *= eth_fraction

    return max(final_pop, 0)

def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=True, sheet_name=sheet_name)
        return output.getvalue()

def create_download_link(file_bytes: bytes, filename: str, label: str):
    """Create full-width HTML download button (without rerun)."""
    b64 = base64.b64encode(file_bytes).decode()
    button_html = f"""<a href="data:application/octet-stream;base64,{b64}" download="{filename}" style="text-decoration:none;">
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
                    box-sizing:border-box;
                    cursor:pointer;
                ">
                {label}
                </div>
            </a>
        """
    st.markdown(button_html, unsafe_allow_html=True)

def create_download_link2(file_bytes: bytes, filename: str, label: str):
    """Create full-width HTML download button (without rerun)."""
    b64 = base64.b64encode(file_bytes).decode()
    button_html = f"""<a href="data:application/octet-stream;base64,{b64}" download="{filename}" style="text-decoration:none;">
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
                    box-sizing:border-box;
                    cursor:pointer;
                ">
                {label}
                </div>
            </a>
        """
    st.markdown(button_html, unsafe_allow_html=True)

def compute_population_coefficients(
    df_ga,
    df_eth,
    region_map,
    gender_selected,
    min_age,
    max_age,
    eth_filter,
    settlement_filter,
    komuna_filter,
    data_collection_method
):
    """
    Kthen koeficientët e popullsisë pas filtrave për:
    - Komunë
    - Regjion
    - Gjinia
    - Grupmosha
    - Vendbanimi
    - Etnia
    Vetëm nëse dimensioni ka ≥ 2 kategori me pop > 0.
    """

    out = []

    # ---------------------------------------------------
    # Prepare df_eth and df_ga based on filters
    # ---------------------------------------------------
    df_pop = df_eth.copy()
    df_pop = df_pop[df_pop["Etnia"].isin(eth_filter)]
    df_pop = df_pop[df_pop["Vendbanimi"].isin(settlement_filter)]
    df_pop = df_pop[df_pop["Komuna"].isin(komuna_filter)]

    if df_pop.empty:
        return pd.DataFrame()

    # df_ga
    df_age = df_ga[df_ga["Komuna"].isin(komuna_filter)]
    df_age = df_age[df_age["Gjinia"].isin(gender_selected)]

    if df_age.empty:
        return pd.DataFrame()

    # Age columns
    age_cols = [c for c in df_age.columns if str(c).isdigit()]

    # max_age automatic for CAWI
    if data_collection_method == "CAWI" and max_age is None:
        max_age = 120

    if max_age is None:
        max_age = max(map(int, age_cols))

    # Filter age columns
    age_mask_cols = [c for c in age_cols if min_age <= int(c) <= max_age]

    df_age["Pop_age"] = df_age[age_mask_cols].sum(axis=1)

    # ---------------------------------------------------
    # Helper to add dimension blocks safely
    # ---------------------------------------------------
    def append_block(dim, pop_series):
        """
        Adds dimension only if ≥2 categories AND total pop > 0.
        Removes categories with Pop=0 before checking.
        """
        # Remove categories with zero population
        pop_series = pop_series[pop_series > 0]

        if len(pop_series) < 2:
            return  # ← SKIP dimension

        total = pop_series.sum()
        if total == 0:
            return  # ← Prevent ZeroDivisionError

        coef_series = pop_series / total

        for cat, pop in pop_series.items():
            out.append({
                "Dimensioni": dim,
                "Kategoria": cat,
                "Populacioni": pop,
                "Pesha": float(coef_series[cat])
            })

    # ---------------------------------------------------
    # 1) Komuna
    # ---------------------------------------------------
    pop_kom = df_pop.groupby("Komuna")["Pop_base"].sum()
    append_block("Komuna", pop_kom)

    # ---------------------------------------------------
    # 2) Regjion
    # ---------------------------------------------------
    df_pop["Regjion"] = df_pop["Komuna"].map(region_map)
    pop_reg = df_pop.groupby("Regjion")["Pop_base"].sum()
    append_block("Regjion", pop_reg)

    # ---------------------------------------------------
    # 3) Etnia
    # ---------------------------------------------------
    pop_eth = df_pop.groupby("Etnia")["Pop_base"].sum()
    append_block("Etnia", pop_eth)

    # ---------------------------------------------------
    # 4) Vendbanimi
    # ---------------------------------------------------
    pop_vb = df_pop.groupby("Vendbanimi")["Pop_base"].sum()
    append_block("Vendbanimi", pop_vb)

    # ---------------------------------------------------
    # 5) Gjinia
    # ---------------------------------------------------
    pop_gender = df_age.groupby("Gjinia")["Pop_age"].sum()
    append_block("Gjinia", pop_gender)

    # ---------------------------------------------------
    # 6) Grupmoshat (dynamic bins)
    # ---------------------------------------------------
    merged_bins, labels = create_dynamic_age_groups(min_age, max_age, data_collection_method)

    # -----------------------------
    # 6) Grupmoshat me etiketa të pastra (18-24, 25-34, 65+)
    # -----------------------------

    merged_bins, labels = create_dynamic_age_groups(min_age, max_age, data_collection_method)

    long_age = []

    for _, row in df_age.iterrows():
        for col in age_mask_cols:
            age = int(col)
            pop = row[col]

            # Gjej bin për këtë moshë
            for idx, (lo, hi) in enumerate(merged_bins):
                if lo <= age <= hi:
                    # Përgatisim formatimin e etiketës
                    if hi >= 85 and data_collection_method!="CAWI":
                        label = f"{lo}+"
                    elif hi >= 65 and data_collection_method=="CAWI":
                        label = f"{lo}+"
                    else:
                        label = f"{lo}-{hi}"
                    long_age.append((label, pop))
                    break

    if long_age:
        df_age_long = pd.DataFrame(long_age, columns=["Age_group", "Count"])
        pop_age_grp = df_age_long.groupby("Age_group")["Count"].sum()

        # Ruaj rendin sipas moshës (p.sh. 18-24 → 25-34 → ... → 65+)
        ordered = sorted(pop_age_grp.index, key=lambda s: int(s.split("-")[0].replace("+","")))
        pop_age_grp = pop_age_grp[ordered]

        append_block("Grupmosha", pop_age_grp)


    # ---------------------------------------------------
    # Return final result
    # ---------------------------------------------------
    return pd.DataFrame(out)

def add_codes_to_coef_df(coef_df, data_collection_method):
    """
    Shton kolonën 'Kodi' në coef_df.
    - Të gjitha dimensionet marrin kodet fikse (si më parë)
    - Vetëm Grupmosha merr kodim dinamik sipas renditjes së saj reale (pas filtrimit)
    """

    # ======================
    # 1. Kodet fikse
    # ======================

    komuna_codes = {
        "Prishtinë":1, "Deçan":2, "Dragash":3, "Ferizaj":4, "Fushë Kosovë":5, 
        "Gjakovë":6, "Gjilan":7, "Gllogoc":8, "Graçanicë":9, "Han i Elezit":10,
        "Istog":11, "Junik":12, "Kaçanik":13, "Kamenicë":14, "Klinë":15,
        "Kllokot":16, "Leposaviq":17, "Leposavic":17, "Lipjan":18, 
        "Malishevë":19, "Mamushë":20, "Mitrovicë":21, "Mitrovica Veriore":22,
        "Novobërdë":23, "Obiliq":24, "Partesh":25, "Pejë":26, "Podujevë":27,
        "Prizren":28, "Rahovec":29, "Ranillug":30, "Skënderaj":31,
        "Suharekë":32, "Shtërpcë":33, "Shtime":34, "Viti":35, "Vushtrri":36,
        "Zubin Potok":37, "Zvecan":38
    }

    region_codes = {
        "Prishtinë":1, "Mitrovicë":2, "Pejë":3, "Prizren":4,
        "Ferizaj":5, "Gjilan":6, "Gjakovë":7
    }

    vb_codes = {"Urban":1, "Rural":2}
    gender_codes = {"Femra":1, "Femer":1, "Mashkull":2, "Meshkuj":2}
    eth_codes = {"Shqiptar":1, "Serb":2, "Tjerë":3, "Tjeter":3}

    # ==========================
    # 2. Fillimisht vendos kodet fikse
    # ==========================

    def get_fixed_code(row):
        dim = row["Dimensioni"]
        cat = row["Kategoria"]

        if dim == "Komuna": return komuna_codes.get(cat)
        if dim == "Regjion": return region_codes.get(cat)
        if dim == "Vendbanimi": return vb_codes.get(cat)
        if dim == "Gjinia": return gender_codes.get(cat)
        if dim == "Etnia": return eth_codes.get(cat)

        # Grupmosha KALO HETU – do mbushet më vonë dinamikisht
        return None

    coef_df["Kodi"] = coef_df.apply(get_fixed_code, axis=1)

    # ==========================
    # 3. KODIMI DINAMIK I GRUPMOSHËS
    # ==========================

    df_age = coef_df[coef_df["Dimensioni"] == "Grupmosha"].copy()

    if not df_age.empty:

        # (a) Parsimi i vlerave të moshës në numra
        parsed = []
        for g in df_age["Kategoria"]:
            g = str(g)

            if "-" in g:
                lo, hi = g.split("-")
                lo, hi = int(lo), int(hi)
            elif g.endswith("+"):
                lo = int(g.replace("+", ""))
                hi = 999
            else:
                continue

            parsed.append((g, lo, hi))

        # (b) Rendit grupmoshat sipas moshës
        parsed_sorted = sorted(parsed, key=lambda x: x[1])

        # (c) Gjenero kodet 1,2,3,... automatikisht
        dynamic_age_codes = {grp: i+1 for i, (grp, _, _) in enumerate(parsed_sorted)}

        # (d) Mbishkruaj kolonën Kodi *vetëm për Grupmoshën*
        coef_df.loc[coef_df["Dimensioni"] == "Grupmosha", "Kodi"] = \
            coef_df.loc[coef_df["Dimensioni"] == "Grupmosha", "Kategoria"].map(dynamic_age_codes)

    return coef_df

def create_dynamic_age_groups(age_min, age_max, data_collection_method):
    """
    Krijon grupmosha dinamike kur përdoruesi ka vendosur max_age.
    Nëse max_age është None → kthen grupmoshat standarde sipas metodës.
    """

    # -----------------------------------------
    # CASE A — Nuk ka max_age → përdor standardet
    # -----------------------------------------
    if age_max is None:
        if data_collection_method == "CAWI":
            base = [(18,24), (25,34), (35,44), (45,54), (55,200)]
        else:
            base = [(18,24), (25,34), (35,44), (45,54), (55,64), (65,200)]

        labels = []
        for lo, hi in base:
            if hi >= 200:
                labels.append(f"{lo}+")
            else:
                labels.append(f"{lo}-{hi}")

        return base, labels

    # -----------------------------------------
    # CASE B — Dynamic bins
    # -----------------------------------------
    base = [
        (18,24),
        (25,34),
        (35,44),
        (45,54),
        (55,64),
        (65,200)
    ]

    if data_collection_method == "CAWI":
        base = [
            (18,24),
            (25,34),
            (35,44),
            (45,54),
            (55,200)
        ]

    hi_age = age_max

    # 1) CLIP bins
    clipped = []
    for lo, hi in base:
        new_lo = max(lo, age_min)
        new_hi = min(hi, hi_age)
        if new_lo <= new_hi:
            clipped.append((new_lo, new_hi))

    # 2) FIX small first bin
    if len(clipped) >= 2:
        lo, hi = clipped[0]
        if (hi - lo + 1) < 5:
            nlo, nhi = clipped[1]
            clipped = [(lo, nhi)] + clipped[2:]

    # 3) FIX small middle bins
    merged = []
    for lo, hi in clipped:
        if merged and (hi - lo + 1) < 5:
            plo, phi = merged[-1]
            merged[-1] = (plo, hi)
        else:
            merged.append((lo, hi))

    # 4) FIX last
    if len(merged) >= 2:
        lo, hi = merged[-1]
        if (hi - lo + 1) < 5:
            plo, phi = merged[-2]
            merged[-2] = (plo, hi)
            merged = merged[:-1]

    # 5) Labels
    labels = []
    for lo, hi in merged:
        if hi >= 200:
            labels.append(f"{lo}+")
        else:
            labels.append(f"{lo}-{hi}")

    return merged, labels

# Load data
try:
    df_eth = load_ethnicity_settlement_data("excel-files/ASK-2024-Komuna-Etnia-Vendbanimi.xlsx")
    df_ga, age_cols = load_gender_age_data("excel-files/ASK-2024-Komuna-Gjinia-Mosha.xlsx")
except Exception as e:
    st.error(f"Gabim gjatë leximit të fajllave: {e}")
    st.stop()

region_map = get_region_mapping()


def generate_spss_syntax_municipality(coef_df, data_collection_method, min_age, max_age):
    """
    Gjeneron SPSS syntax për municipality me:
    - Grupmosha standarde nëse max_age = None
    - Grupmosha dinamike kur përdoruesi zgjedh max_age
    """

    out = "* Encoding: UTF-8.\n\n"

    # ===============================================================
    # 1) GJENERO LISTËN E GRUPMOSHAVE
    # ===============================================================

    df_age = coef_df[coef_df["Dimensioni"] == "Grupmosha"].copy()

    # ---------------------------------------------------
    # CASE 1 — Nuk ka max_age → përdor grupmosha standarde
    # ---------------------------------------------------
    if max_age is None:
        if data_collection_method == "CAWI":
            groups = [
                ("18", "24"),
                ("25", "34"),
                ("35", "44"),
                ("45", "54"),
                ("55", "HI")
            ]
            labels = ["18 - 24", "25 - 34", "35 - 44", "45 - 54", "55+"]
        else:
            groups = [
                ("18", "24"),
                ("25", "34"),
                ("35", "44"),
                ("45", "54"),
                ("55", "64"),
                ("65", "HI")
            ]
            labels = ["18 - 24", "25 - 34", "35 - 44", "45 - 54", "55 - 64", "65+"]

        # KODIMET
        recode_lines = [f"(LO THRU {hi} = {i+1})" for i, (lo, hi) in enumerate(groups)]
        value_labels = "\n".join([f" {i+1} '{lbl}'" for i, lbl in enumerate(labels)])

    else:
        # ---------------------------------------------------
        # CASE 2 — max_age ekziston → përdor grupmosha dinamike
        # ---------------------------------------------------

        def parse_range(g):
            g = g.strip()
            if "+" in g:
                lo = int(g.replace("+", "").strip())
                hi = 999
                return lo, hi

            lo, hi = g.replace(" ", "").split("-")
            lo = int(lo)
            hi = int(hi)

            # Çdo grup që përfundon me >85 e trajtojmë si open-ended
            if hi >= 85:
                return lo, 999

            return lo, hi

        df_age["lo_hi"] = df_age["Kategoria"].apply(parse_range)
        df_age = df_age.sort_values(by="lo_hi", key=lambda x: x.apply(lambda y: y[0]))

        recode_lines = []
        value_labels = ""
        code = 1

        for _, row in df_age.iterrows():
            lo, hi = row["lo_hi"]

            if hi == 999:
                recode_lines.append(f"(LO THRU HI = {code})")
                value_labels += f" {code} '{lo}+' \n"
            else:
                recode_lines.append(f"(LO THRU {hi} = {code})")
                value_labels += f" {code} '{lo}-{hi}' \n"

            code += 1

    # ===============================================================
    # 2) VIZUAL BINNING
    # ===============================================================
    out += (
        "* Visual Binning.\n"
        "* Mosha.\n"
        "RECODE D2 (MISSING=COPY)\n    "
        + "\n    ".join(recode_lines)
        + "\n (ELSE=SYSMIS) INTO Grupmoshat.\n"
        "VARIABLE LABELS Grupmoshat 'Mosha (Binned)'.\n"
        "FORMATS Grupmoshat (F5.0).\n"
        "VALUE LABELS Grupmoshat\n"
        f"{value_labels}"
        "VARIABLE LEVEL Grupmoshat (ORDINAL).\n"
        "EXECUTE.\n\n"
    )

    # ===============================================================
    # 3) SPSSINC RAKE
    # ===============================================================
    out += "SPSSINC RAKE\n"

    dim_order = ["Gjinia", "Grupmosha", "Vendbanimi", "Etnia"]
    dim_i = 1

    for dim in dim_order:
        df_dim = coef_df[coef_df["Dimensioni"] == dim]
        if df_dim.empty:
            continue

        out += f"DIM{dim_i}={dim} "
        for _, row in df_dim.iterrows():
            out += f"{int(row['Kodi'])} {row['Pesha']}\n"

        dim_i += 1

    out += "FINALWEIGHT=peshat.\n"

    return out
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
    step=100
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
    # CALCULATE MARGIN OF ERROR FOR THE MUNICIPALITY
    # =====================================================

    z = 1.96    # 95% confidence
    p = 0.5     # worst-case proportion
    n = N       # sample size for municipality
    Npop = total_pop

    if Npop > n:
        fpc = ((Npop - n) / (Npop - 1)) ** 0.5   # finite population correction
    else:
        fpc = 1.0

    moe = z * ((p * (1 - p)) / n) ** 0.5 * fpc
    moe_percent = moe * 100


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

        # EXCEPTION: For Junik → no minimum threshold required
        if komuna == "Junik":
            weights = df_rural["PopFilt"] / df_rural["PopFilt"].sum()
            floats = weights * rural_n
            alloc = controlled_rounding(floats, rural_n)
            df_rural["Intervista"] = alloc
            return df_rural

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
    
    # Përgatit tekstin për gjininë
    if len(gender_selected) == 1:
        gender_text = f"{gender_selected[0]}"
    else:
        gender_text = "Femra, Meshkuj"

    if len(eth_filter) == 1 or len(eth_filter) == 2:
        ethnicity_text = ", ".join(eth_filter)
    else:
        ethnicity_text = "Shqiptar, Serb, Tjerë"

    def load_svg(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()

    icon_sample = load_svg("images/sample-people.svg")
    icon_strata = load_svg("images/strata.svg")
    icon_demo = load_svg("images/demographics.svg")


    col1, col2 = st.columns(2)

    with col1:
        with st.container():
            st.markdown(f"""
            <div class='card'>
                <div class='card-title'>
                    {icon_sample} Mostra
                </div>
                <div class='card-value'>Totali i mostrës: <b>{N}</b></div>
                <div class='card-value'>Marzha e gabimit: <b>± {moe_percent:.2f}%</b></div>
                <div class='card-value'>Intervali i besimit: <b>95%</b></div>
            </div>
            """, unsafe_allow_html=True)

    with col2:
        with st.container():
            st.markdown(f"""
            <div class='card'>
                <div class='card-title'>
                        {icon_demo} Demografia
                </div>
                <div class='card-value'>Grupmosha: <b>{age_text}</b></div>
                <div class='card-value'>Gjinia: <b>{gender_text}</b></div>
                <div class='card-value'>Etnia: <b>{ethnicity_text}</b></div>
            </div>
            """, unsafe_allow_html=True)

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
    deck_html = deck.to_html(as_string=True)
    html_bytes = deck_html.encode("utf-8")
    # Butoni i shkarkimit
    create_download_link(html_bytes, "psu_map.html", "Shkarko hartën (HTML)")
    
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
    data_collection_method=data_collection_method
)
    # ============================================
    # FIX: Formatimi i etiketimeve të Grupmoshave
    # ============================================
    def fix_age_label(label):
        label = str(label).strip()

        # Shembuj: "18-24", "65-200", "65-120", "55-64"
        if "-" in label:
            lo, hi = label.split("-")
            lo = lo.strip()
            hi = hi.strip()

            # HI → kthente 200 → duhet të bëhet "+"
            if hi in ["200", "300", "999"]:
                return f"{lo}+"

            # Përndryshe normalizim "18-24"
            return f"{lo}-{hi}"

        # Për raste si "65+"
        if label.endswith("+"):
            return label

        return label

    # Apliko fix vetëm te Grupmosha
    coef_df.loc[coef_df["Dimensioni"] == "Grupmosha", "Kategoria"] = \
        coef_df.loc[coef_df["Dimensioni"] == "Grupmosha", "Kategoria"].apply(fix_age_label)


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
    data_collection_method, 
    min_age,
    max_age
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
