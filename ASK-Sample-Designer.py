import streamlit as st
import pandas as pd
import numpy as np

# =========================
# CONFIG
# =========================

st.set_page_config(
    page_title="Dizajnimi i Mostrës Nacionale",
    layout="wide"
)

# =========================
# HELPERS
# =========================

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
        "Fushë Kosova": "Prishtinë",
        "Gjakovë": "Gjakovë",
        "Gjilan": "Gjilan",
        "Gllogoc": "Prishtinë",
        "Graçanicë": "Prishtinë",
        "Hani I Elezit": "Ferizaj",
        "Istog": "Pejë",
        "Junik": "Gjakovë",
        "Kaçanik": "Ferizaj",
        "Kamenicë": "Gjilan",
        "Klinë": "Pejë",
        "Kllokot": "Gjilan",
        "Leposaviq": "Mitrovicë",
        "Lipjan": "Prishtinë",
        "Malishevë": "Prizren",
        "Mamushë": "Prizren",
        "Mitrovicë": "Mitrovicë",
        "Mitrovica Veriut": "Mitrovicë",
        "Novobërdë": "Gjilan",
        "Obiliq": "Prishtinë",
        "Partesh": "Gjilan",
        "Peja": "Pejë",
        "Podujeva": "Prishtinë",
        "Prishtina": "Prishtinë",
        "Prizren": "Prizren",
        "Rahovec": "Gjakovë",
        "Ranillug": "Gjilan",
        "Shtërpcë": "Ferizaj",
        "Shtime": "Ferizaj",
        "Skenderaj": "Mitrovicë",
        "Suharekë": "Prizren",
        "Viti": "Gjilan",
        "Vushtrri": "Mitrovicë",
        "Zubin Potok": "Mitrovicë",
        "Zveçan": "Mitrovicë"
    }
    return region_map

def compute_gender_age_coefficients(df_ga: pd.DataFrame,
                                    age_cols,
                                    selected_genders,
                                    min_age: int,
                                    max_age: int | None) -> pd.Series:
    """
    Compute coefficient per Komuna:
    coef(komuna) = Pop_selected(komuna) / Pop_total(komuna)

    - Pop_selected: individuals that match selected genders & [min_age, max_age]
    - Pop_total: all genders, all ages
    """
    df = df_ga.copy()

    # Sort age columns numerically
    age_cols_sorted = sorted(age_cols, key=lambda x: int(str(x)))
    max_available_age = int(str(age_cols_sorted[-1]))

    if max_age is None:
        max_age = max_available_age

    min_age = int(min_age)
    max_age = int(max_age)

    # Selected ages
    selected_age_cols = [
        c for c in age_cols_sorted
        if min_age <= int(str(c)) <= max_age
    ]

    if not selected_age_cols:
        # No matching age columns: coefficient 0 for all
        return pd.Series(
            0.0,
            index=df["Komuna"].unique()
        )

    # Total population (all genders, all ages)
    df["Pop_all_ages"] = df[age_cols_sorted].sum(axis=1)

    # Filtered population by age range
    df["Pop_age_range"] = df[selected_age_cols].sum(axis=1)

    def calc_coef(group: pd.DataFrame) -> float:
        total_pop = group["Pop_all_ages"].sum()

        if total_pop == 0:
            return 0.0

        # If no gender specified (safety) -> use all genders
        if not selected_genders:
            num = group["Pop_age_range"].sum()
        else:
            num = group.loc[
                group["Gjinia"].isin(selected_genders),
                "Pop_age_range"
            ].sum()

        if num <= 0:
            return 0.0

        return float(num) / float(total_pop)

    coef_by_komuna = df.groupby("Komuna").apply(calc_coef)
    return coef_by_komuna

def compute_age_os_share(df_ga: pd.DataFrame,
                         age_cols,
                         selected_genders,
                         global_min_age: int,
                         global_max_age: int | None,
                         os_min_age: int,
                         os_max_age: int) -> pd.Series:
    """
    Kthen për secilën Komunë:
    share_OS(komuna) = Pop( OS_min–OS_max ) / Pop( global_min–global_max )

    – përdor vetëm gjinitë e zgjedhura te `selected_genders`
    – nëse s'ka popullsi në intervalin global -> 0
    """

    df = df_ga.copy()

    # Sorto kolonat e moshës
    age_cols_sorted = sorted(age_cols, key=lambda x: int(str(x)))
    max_available_age = int(str(age_cols_sorted[-1]))

    # Global max nëse është None
    if global_max_age is None:
        global_max_age = max_available_age

    # Konverto në int
    global_min_age = int(global_min_age)
    global_max_age = int(global_max_age)
    os_min_age = int(os_min_age)
    os_max_age = int(os_max_age)

    # Kufizo OS brenda intervalit global
    os_min_age = max(os_min_age, global_min_age)
    os_max_age = min(os_max_age, global_max_age)

    if os_min_age > os_max_age:
        # Interval OS bosh -> gjithmonë 0
        return pd.Series(0.0, index=df["Komuna"].unique())

    global_cols = [
        c for c in age_cols_sorted
        if global_min_age <= int(str(c)) <= global_max_age
    ]
    os_cols = [
        c for c in age_cols_sorted
        if os_min_age <= int(str(c)) <= os_max_age
    ]

    df["Pop_global"] = df[global_cols].sum(axis=1)
    df["Pop_os"] = df[os_cols].sum(axis=1)

    def agg(group: pd.DataFrame) -> float:
        if selected_genders:
            group = group[group["Gjinia"].isin(selected_genders)]

        pop_global = group["Pop_global"].sum()
        if pop_global <= 0:
            return 0.0

        pop_os = group["Pop_os"].sum()
        if pop_os <= 0:
            return 0.0

        return float(pop_os) / float(pop_global)

    share_by_komuna = df.groupby("Komuna").apply(agg)
    return share_by_komuna

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
   
# Helper to identify strata belonging to each oversample variable
def mask_for_oversample(grouped, variable, params):

    if variable == "Komuna":
        return grouped[base_col] == params["value"]

    if variable == "Regjion":
        return grouped[base_col] == params["value"]

    if variable == "Vendbanimi":
        return grouped["Sub"].str.endswith(params["value"])

    if variable == "Etnia":
        return grouped["Sub"].str.startswith(params["value"])

    if variable == "Gjinia":
        if "Gjinia" in grouped.columns:
            return grouped["Gjinia"] == params["value"]
        else:
            return pd.Series(False, index=grouped.index)

    if variable == "Mosha":
        # "OS" është segmenti i krijuar më herët
        if "AgeSeg" in grouped.columns:
            return grouped["AgeSeg"] == "OS"
        return pd.Series(False, index=grouped.index)

    return pd.Series(False, index=grouped.index)

# Load data
try:
    df_eth = load_ethnicity_settlement_data("ASK-2024-Komuna-Etnia-Vendbanimi.xlsx")
    df_ga, age_cols = load_gender_age_data("ASK-2024-Komuna-Gjinia-Mosha.xlsx")
except Exception as e:
    st.error(f"Gabim gjatë leximit të fajllave: {e}")
    st.stop()

region_map = get_region_mapping()

# =========================
# UI: SIDEBAR
# =========================

st.title("Dizajnimi i Mostrës Nacionale")

st.sidebar.header("Parametrat kryesorë")

# Total sample size
n_total = st.sidebar.number_input(
    "Numri total i mostrës (n)",
    min_value=1,
    value=1065,
    step=10
)

# Primary stratification
primary_level = st.sidebar.selectbox(
    "Ndarja kryesore",
    options=["Komunë", "Regjion"],
    index=0
)

# Sub-stratification (can choose Vendbanim, Etnia, or both)
sub_options = st.sidebar.multiselect(
    "Nën-ndarja (mund të zgjedhësh një ose të dyja)",
    options=["Vendbanim", "Etnia"],
    default=["Vendbanim", "Etnia"]
)

st.sidebar.markdown("---")

# Demographic filters
st.sidebar.subheader("Filtrat demografikë")

# Komuna filter
komuna_filter = st.sidebar.multiselect(
    "Komunat që përfshihen",
    options=sorted(df_eth["Komuna"].unique()),
    default=sorted(df_eth["Komuna"].unique())
)

gender_selected = st.sidebar.multiselect(
    "Gjinia për përfshirje në mostër",
    options=["Meshkuj", "Femra"],
    default=["Meshkuj", "Femra"]
)

min_age = st.sidebar.number_input(
    "Mosha minimale (obligative)",
    min_value=0,
    value=18,
    step=1
)

max_age = st.sidebar.text_input(
    "Mosha maksimale (opsionale — lëre bosh nëse nuk ka kufi)"
)

max_age = int(max_age) if max_age.strip() else None

# Ethnicity filter (these also act as possible sub-dimensions if Etnia selected)
eth_filter = st.sidebar.multiselect(
    "Etnitë që përfshihen",
    options=["Shqiptar", "Serb", "Tjerë"],
    default=["Shqiptar", "Serb", "Tjerë"]
)

# Settlement filter
settlement_filter = st.sidebar.multiselect(
    "Vendbanimi që përfshihet",
    options=["Urban", "Rural"],
    default=["Urban", "Rural"]

)
# Oversampling
st.sidebar.markdown("---")

oversample_enabled = st.sidebar.checkbox("Oversample", value=False)

oversample_inputs = {}

if oversample_enabled:

    oversample_vars = st.sidebar.multiselect(
        "Zgjidh deri në 2 variabla për oversample:",
        options=["Regjion", "Komuna", "Vendbanimi", "Gjinia", "Etnia", "Mosha"],
        max_selections=2
    )

    for var in oversample_vars:
        st.sidebar.markdown(f"**{var}**")

        if var == "Mosha":
            min_over_age = st.sidebar.number_input(
                f"Grupmosha minimale ({var})",
                min_value=0, value=18, step=1,
                key=f"min_{var}"
            )
            max_over_age = st.sidebar.number_input(
                f"Grupmosha maksimale ({var})",
                min_value=min_over_age, value=24, step=1,
                key=f"max_{var}"
            )
            oversample_n = st.sidebar.number_input(
                f"Numri i anketave për këtë grupmoshë",
                min_value=1, value=50, step=10,
                key=f"n_{var}"
            )

            oversample_inputs[var] = {
                "min_age": min_over_age,
                "max_age": max_over_age,
                "n": oversample_n
            }

        else:
            df_tmp = df_eth[
                (df_eth["Etnia"].isin(eth_filter)) &
                (df_eth["Vendbanimi"].isin(settlement_filter)) &
                (df_eth["Komuna"].isin(komuna_filter))
            ]
            
            if var == "Komuna":
                # kufizojmë komunat që kanë popullsi > 0 pas filtrave bazë
                valid_kom = df_tmp.groupby("Komuna")["Pop_base"].sum()
                valid_kom = valid_kom[valid_kom > 0].index.tolist()

                options = sorted(valid_kom)


            elif var == "Regjion":
                valid_kom = df_tmp.groupby("Komuna")["Pop_base"].sum()
                valid_kom = valid_kom[valid_kom > 0].index.tolist()

                valid_reg = {region_map[k] for k in valid_kom if k in region_map}

                options = sorted(valid_reg)


            elif var == "Vendbanimi":
                options = sorted(settlement_filter) 

            elif var == "Etnia":
                options = sorted(eth_filter)          

            elif var == "Gjinia":
                options = sorted(gender_selected)

            elif var == "Mosha":
                # Zgjedhja për moshën nuk kufizohet
                options = []

            else:
                options = []

            val = st.sidebar.selectbox(
                f"Njësia nga {var} që do të jetë oversample",
                options=options,
                key=f"val_{var}"
            )
            oversample_n = st.sidebar.number_input(
                f"Numri i anketave për {var} = {val}",
                min_value=1, value=50, step=10,
                key=f"n_{var}"
            )

            oversample_inputs[var] = {
                "value": val,
                "n": oversample_n
            }

st.sidebar.markdown("---")

seed = 42

st.sidebar.markdown("Kliko më poshtë për të llogaritur shpërndarjen.")
run_button = st.sidebar.button("Gjenero shpërndarjen e mostrës")

# =========================
# MAIN LOGIC
# =========================

if run_button:

    # 1) Filter ethnicity & settlement (these are demographic filters)
    df = df_eth.copy()
    df = df[df["Etnia"].isin(eth_filter)]
    df = df[df["Vendbanimi"].isin(settlement_filter)]
    df = df[df["Komuna"].isin(komuna_filter)]

    if df.empty:
        st.error("Asnjë kombinim nuk përputhet me filtrat e zgjedhur (Etnia/Vendbanimi).")
        st.stop()

    # 2) Compute gender+age coefficients per Komuna
    coef_by_komuna = compute_gender_age_coefficients(
        df_ga,
        age_cols=age_cols,
        selected_genders=gender_selected,
        min_age=min_age,
        max_age=max_age
    )

    # Mapojmë gjininë origjinale në df_eth duke përdorur df_ga si referencë
    gender_map = df_ga.groupby("Komuna")["Gjinia"].apply(list).to_dict()

    df["Gjinia"] = df["Komuna"].map(lambda k: gender_map.get(k, ["Meshkuj","Femra"]))
                    
    # zgjerim i rreshtave për çdo gjini (ndryshe OS nuk punon)
    df = df.explode("Gjinia")

    # Attach coefficient to df (missing komuna -> coef 0)
    df["coef_gender_age"] = df["Komuna"].map(coef_by_komuna).fillna(0.0)

    # 3) Adjusted population for filters (Etnia, Vendbanim, Gjinia, Mosha)
    df["Pop_adj"] = df["Pop_base"] * df["coef_gender_age"]

    # Remove rows with zero adjusted population
    df = df[df["Pop_adj"] > 0]

    if df.empty:
        st.error("Pas aplikimit të koeficientëve (gjinia/mosha), Pop_adj është 0 për të gjitha njësitë.")
        st.stop()

    # 4) Primary stratification
    if primary_level == "Regjion":
        if not region_map:
            st.warning(
                "Ndarja sipas Regjionit kërkon të plotësohet 'region_map' në kod. "
                "Aktualisht nuk ka mapping, prandaj po vazhdohet vetëm me nivel Komune."
            )
            base_col = "Komuna"
        else:
            df["Regjion"] = df["Komuna"].map(region_map)
            df = df.dropna(subset=["Regjion"])
            base_col = "Regjion"
    else:
        base_col = "Komuna"

    # 5) Sub-stratification labels
    # Ensure consistent sorting at ethnicity
    eth_order = ["Shqiptar", "Serb", "Tjerë"]
    df["Etnia"] = pd.Categorical(df["Etnia"], categories=eth_order, ordered=True)

    # Combine ethnicity with settlement
    if "Etnia" in sub_options and "Vendbanim" in sub_options:
        df["Sub"] = df["Etnia"].astype(str) + " - " + df["Vendbanimi"].astype(str)
    elif "Etnia" in sub_options:
        df["Sub"] = df["Etnia"].astype(str)
    elif "Vendbanim" in sub_options:
        df["Sub"] = df["Vendbanimi"].astype(str)
    else:
        df["Sub"] = "Total"

    grouped = (
        df.groupby([base_col, "Sub"], as_index=False)["Pop_adj"]
        .sum()
        .rename(columns={"Pop_adj": "Pop_stratum"})
    )

    grouped = grouped.reset_index(drop=True)

    precomputed_masks = {}
    for var, params in oversample_inputs.items():
        precomputed_masks[var] = mask_for_oversample(grouped, var, params)

    # Sort columns
    sub_order = []
    for eth in eth_order:
        for vb in ["Urban", "Rural"]:
            sub_order.append(f"{eth} - {vb}")
    sub_order += eth_order + ["Urban", "Rural", "Total"]  # për raste të tjera

    # Filtwe columns
    existing_subs = sorted(grouped["Sub"].unique(), key=lambda x: sub_order.index(x) if x in sub_order else 999)
    grouped["Sub"] = pd.Categorical(grouped["Sub"], categories=existing_subs, ordered=True)


    total_pop = grouped["Pop_stratum"].sum()
    if total_pop <= 0:
        st.error("Popullsia totale pas filtrave është 0. Nuk mund të alokohet mostra.")
        st.stop()

    # ================================
    # 7a) Oversampling
    # ================================

    grouped["n_alloc"] = 0

    oversample_items = list(oversample_inputs.items())

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

    # 0) Nuk ka oversample fare → alokim i thjeshtë proporcional
    if not oversample_items:
        total_pop = grouped["Pop_stratum"].sum()
        weights = grouped["Pop_stratum"] / total_pop
        floats = weights * n_total
        grouped["n_alloc"] = controlled_rounding(floats.to_numpy(), n_total, seed)

    # 1) Vetëm një oversample
    elif len(oversample_items) == 1:
        varA, paramsA = oversample_items[0]
        nA = int(paramsA["n"])

        maskA = mask_for_oversample(grouped, varA, paramsA)

        # Alokim i OS‐it në maskA
        alloc_A = alloc_to_mask(maskA, nA)
        used_A = int(alloc_A.sum())

        # Pjesa e mbetur shkon në stratum‐et jashtë OS
        remaining = n_total - used_A
        if remaining < 0:
            remaining = 0

        mask_rest = ~maskA
        alloc_rest = alloc_to_mask(mask_rest, remaining)

        grouped["n_alloc"] = alloc_A + alloc_rest

    # 2) Dy variabla oversample (rasti i overlapp‐it)
    else:
        # =========================================
        # PATCH 3 — Oversample me dy variabla
        # =========================================

        (varA, paramsA), (varB, paramsB) = oversample_items[:2]
        nA = int(paramsA["n"])
        nB = int(paramsB["n"])

        maskA = precomputed_masks[varA]
        maskB = precomputed_masks[varB]

        # 1) Shpërndaj OS_B (p.sh. Rural = 800)
        alloc_B = alloc_to_mask(maskB, nB)

        # 2) Overlap: pjesa që OS_B jep brenda OS_A
        overlap_mask = maskA & maskB
        overlap_from_B = int(alloc_B[overlap_mask].sum())

        # 3) Pjesa që i mungon OS_A mbi overlapp
        remaining_A = max(nA - overlap_from_B, 0)
        alloc_A_extra = alloc_to_mask(maskA & ~maskB, remaining_A)

        # 4) Union i dy OS-ve
        total_os_used = int(alloc_B.sum() + alloc_A_extra.sum())

        remaining = max(n_total - total_os_used, 0)
        alloc_rest = alloc_to_mask(~(maskA | maskB), remaining)

        grouped["n_alloc"] = alloc_B + alloc_A_extra + alloc_rest

    # ================================
    # 8) Heq kolonat që s'duhen para pivot-it
    # ================================
    drop_cols = []
    if "Gjinia" in grouped.columns:
        drop_cols.append("Gjinia")
    if "AgeSeg" in grouped.columns:
        drop_cols.append("AgeSeg")

    if drop_cols:
        grouped = (
            grouped.groupby([base_col, "Sub"])[["Pop_stratum", "n_alloc"]]
                .sum()
                .reset_index()
        )

    # 9) Prepare pivot table: rows = primary, columns = sub-dimensions
    pivot = grouped.pivot(
        index=base_col,
        columns="Sub",
        values="n_alloc"
    ).fillna(0).astype(int)

    # Add total per primary
    pivot["Total"] = pivot.sum(axis=1)
    
    # ==========================================================
    #  OVERSAMPLE GENDER/MOSHA pas pivot (në nivel KOMUNE)
    # ==========================================================

     # Marrim TOTAL-in e komunës nga pivot
    pivot_totals = pivot["Total"].copy()

        # -------------------------
        # 1) GJINIA
        # -------------------------
    if "Gjinia" in oversample_inputs:

        os_gender = oversample_inputs["Gjinia"]["value"]
        os_n = int(oversample_inputs["Gjinia"]["n"])

        if os_n > n_total:
            st.warning(
            f"Vërejtje: Kuota e alokuar ({os_n}) për oversample tek ({os_gender}) është më e madhe se N = ({n_total}). "
            "Shëno një kuotë tjetër për oversample."
        )

        # Popullsia sipas gjinisë per komunë
        pop_by_gender = (
            df_ga.groupby(["Komuna", "Gjinia"])[age_cols]
            .sum()
            .sum(axis=1)
            .unstack(fill_value=0)
            .reindex(pivot.index)
            .fillna(0)
            )

        pop_os = pop_by_gender[os_gender]
        weight_os = pop_os / pop_os.sum()

        # Alokimi për OS
        os_alloc = (weight_os * os_n).round().astype(int)

        # Alokimi për pjesën tjetër
        leftover = pivot_totals - os_alloc

        if os_gender == "Femra":
            pivot["Femra"] = controlled_rounding(os_alloc, os_n)
            pivot["Meshkuj"] = pivot_totals - pivot["Femra"]
        else:
            pivot["Meshkuj"] = controlled_rounding(os_alloc, os_n)
            pivot["Femra"] = pivot_totals - pivot["Meshkuj"]

        # -------------------------
        # 2) MOSHA
        # -------------------------
    if "Mosha" in oversample_inputs:

        params_age = oversample_inputs["Mosha"]
        os_min = params_age["min_age"]
        os_max = params_age["max_age"]
        os_n = int(params_age["n"])

        if os_n > n_total:
            st.warning(
            f"Vërejtje: Kuota e alokuar ({os_n}) për oversample tek ({os_min}-{os_max}) është më e madhe se N = ({n_total}). "
            "Shëno një kuotë tjetër për oversample."
        )

        # Lista e moshave që ekzistojnë në dataset
        age_cols_sorted = sorted(age_cols, key=lambda x: int(str(x)))

        # Grupi OS (18–30 p.sh.)
        range_os = [c for c in age_cols_sorted if os_min <= int(c) <= os_max]

        # Grupi jashtë OS (31+ p.sh.)
        range_non = [c for c in age_cols_sorted if int(c) > os_max]

        # Popullsia sipas moshës per komunë
        pop_by_age = (
            df_ga.groupby("Komuna")[age_cols_sorted]
            .sum()
            .reindex(pivot.index)
            .fillna(0)
            )

        pop_os = pop_by_age[range_os].sum(axis=1)
        pop_non = pop_by_age[range_non].sum(axis=1)

        # Pesha për OS
        weight_os_age = pop_os / pop_os.sum()

        # 1) SHPËRNDA OS për moshë
        os_alloc_age = (weight_os_age * os_n).round().astype(int)

        # 2) Alokimi final për moshë
        age_label_os = f"{os_min}–{os_max}"
        age_label_non = f"{os_max+1}+"

        pivot[age_label_os] = controlled_rounding(os_alloc_age, os_n)
        pivot[age_label_non] = pivot_totals - pivot[age_label_os]

    # Safety: ensure global total matches n_total
    global_total = int(pivot["Total"].sum())

    pivot.loc["Total"] = pivot.sum(numeric_only=True)

    st.subheader("Tabela e ndarjes së mostrës")

    # Përgatit tekstin për grupmoshën
    if max_age is None:
        age_text = f"Grupmosha: {min_age}+"
    else:
        age_text = f"Grupmosha: {min_age}–{max_age}"

    # Përgatit tekstin për gjininë
    if len(gender_selected) == 1:
        gender_text = f"Gjinia: {gender_selected[0]}"
    else:
        gender_text = ""

    # Linja kryesore
    caption_main = (
        f"Ndarja kryesore: **{primary_level}** | "
        f"Nën-ndarja: **{', '.join(sub_options) if sub_options else 'Asnjë'}** | "
        f"Totali i mostrës: **{n_total}** | "
        f"Totali i alokuar: **{global_total}**"
    )

    # Linja shtesë për filtrat demografikë
    caption_extra = " | ".join(filter(None, [age_text, gender_text]))

    # Shfaq të dyja linjat
    st.caption(caption_main)
    if caption_extra:
        st.caption(caption_extra)

    st.dataframe(pivot, use_container_width=True)

    if global_total != n_total:
        st.warning(
            f"Vërejtje: Totali i alokuar ({global_total}) nuk përputhet me N = ({n_total}). "
            "Kontrollo koeficientët dhe numerikën."
        )

    # 10) Show long format result (optional, më teknik)
    with st.expander("Shfaq tabelën e plotë të stratum-eve (long format)", expanded=False):
        display_cols = [base_col, "Sub", "Pop_stratum", "n_alloc"]
        st.dataframe(grouped[display_cols], use_container_width=True)

    # 11) Download buttons (secila tabelë veç e veç në Excel)

    def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=True, sheet_name=sheet_name)
        return output.getvalue()

    import base64

    def create_download_link(file_bytes: bytes, filename: str, label: str):
        """Create full-width HTML download button (without rerun)."""
        b64 = base64.b64encode(file_bytes).decode()
        button_html = f"""
            <a href="data:application/octet-stream;base64,{b64}" download="{filename}" style="text-decoration:none;">
                <div style="
                    background-color:#0054a3;
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


    # 📘 Pivot table (Excel)
    pivot_excel = df_to_excel_bytes(pivot, sheet_name="Mostra")
    create_download_link(
        file_bytes=pivot_excel,
        filename="mostra_e_gjeneruar.xlsx",
        label="Shkarko Mostrën"
    )

    # 📘 Strata table (Excel)
    strata_excel = df_to_excel_bytes(grouped, sheet_name="Strata")
    create_download_link(
        file_bytes=strata_excel,
        filename="mostra_strata.xlsx",
        label="Shkarko Strata"
    )

else:
    st.info("Cakto parametrat kryesorë dhe kliko **'Gjenero shpërndarjen e mostrës'** për të dizajnuar mostrën.")
