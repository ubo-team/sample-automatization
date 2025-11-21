import streamlit as st
import pandas as pd
import numpy as np
import re

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

def fix_minimum_allocations(
    pivot: pd.DataFrame,
    df_eth: pd.DataFrame,
    region_map: dict,
    min_total: int = 3,
    min_eth: int = 3,      # threshold for removing (total eth < 3)
    min_vb: int = 2        # not used for ethnicity removal now, only for settlement logic
) -> pd.DataFrame:

    pivot_fixed = pivot.copy()
    municipalities = list(pivot_fixed.index)

    # region lookup
    region_of = pivot_fixed.index.to_series().map(region_map)

    # store initial totals for receiver limit
    initial_total = pivot_fixed["Total"].copy()

    # identify ethnicity columns (Shqiptar, Serb, Tjeter)
    eth_cols = [c for c in pivot_fixed.columns if any(x in c for x in ["Shqiptar", "Serb", "Tjerë"])]

    # map majority ethnicity by true population
    def compute_majority(kom):
        subset = df_eth[df_eth["Komuna"] == kom]
        s = subset.groupby("Etnia")["Pop_base"].sum()
        if s.empty:
            return None
        return s.idxmax()

    majority = {k: compute_majority(k) for k in municipalities}

    # allowed matrix (columns that existed initially)
    allowed = (pivot > 0)

    # helper — find receivers for ethnicity
    def receiver_candidates(eth, col, kom):
        receivers = []
        for r in municipalities:
            if r == kom:
                continue

            # SERB and SHQIPTAR have strict rules
            if eth in ["Serb", "Shqiptar"]:
                if majority[r] != eth:
                    continue
                if not allowed.at[r, col]:
                    continue

            # Tjetër has no majority or allowed restriction
            if eth == "Tjerë":
                if not allowed.at[r, col]:
                    continue

            # receiver limit check
            if pivot_fixed.at[r, "Total"] >= initial_total[r] + 3:
                continue

            receivers.append(r)

        # region-first
        in_region = [r for r in receivers if region_of[r] == region_of[kom]]
        if in_region:
            return in_region
        
        return receivers

    # -----------------------------------------------------
    # ETHNIC REALLOCATION (core logic)
    # -----------------------------------------------------
    # We keep Urban/Rural separately, but remove all
    # units for an ethnicity if total < 3
    # -----------------------------------------------------
    ethnic_groups = {
        "Shqiptar": [c for c in eth_cols if c.startswith("Shqiptar")],
        "Serb":     [c for c in eth_cols if c.startswith("Serb")],
        "Tjerë":   [c for c in eth_cols if c.startswith("Tjerë")]
    }

    for kom in municipalities:
        if kom == "Total":
            continue

        for eth, cols in ethnic_groups.items():

            # total across Urban/Rural
            total_eth = sum(pivot_fixed.at[kom, c] for c in cols)

            # OK if >= 3
            if total_eth >= 3:
                continue

            # nothing to remove if 0
            if total_eth == 0:
                continue

            # number of units to remove = all units
            units_to_move = total_eth

            # move Urban first then Rural (or reverse)
            for col in cols:
                while pivot_fixed.at[kom, col] > 0:

                    # find receivers
                    recv_list = receiver_candidates(eth, col, kom)

                    if not recv_list:
                        # no available receivers → stop trying
                        print(f"[WARNING] No receivers found for {eth} / {col} from {kom}")
                        break

                    recv = recv_list[0]

                    # transfer 1 unit FROM kom TO recv
                    pivot_fixed.at[kom, col] -= 1
                    pivot_fixed.at[recv, col] += 1

                    pivot_fixed.at[kom, "Total"] -= 1
                    pivot_fixed.at[recv, "Total"] += 1

    # -----------------------------------------------------
    # RECOMPUTE TOTALS
    # -----------------------------------------------------

    subs = [c for c in pivot_fixed.columns if c not in ["Total"]]
    pivot_fixed["Total"] = pivot_fixed[subs].sum(axis=1)

    # Ribëj totalin e fundit
    pivot_fixed.loc["Total"] = pivot_fixed.sum(numeric_only=True)

    return pivot_fixed

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

def compute_num_psu(total_interviews: int, k: int):
    """
    Rregulli yt:
    - q = T // k, r = T % k
    - nëse r == 0 → PSU të plota
    - nëse r <= k/2 → nuk shtohet PSU, leftover = r (shpërndahet te PSU-të më të mëdha)
    - nëse r > k/2 → shtohet një PSU shtesë me madhësi r
    """
    if total_interviews <= 0:
        return 0, 0, 0

    q = total_interviews // k
    r = total_interviews % k
    half_k = k / 2.0

    if r == 0:
        return q, 0, 0

    if r <= half_k:
        return q, r, 0
    else:
        return q + 1, 0, r


def select_psus_for_municipality(
    komuna: str,
    total_interviews: int,
    df_psu_mun: pd.DataFrame,
    k: int,
    required_ethnicities: list[str]
) -> pd.DataFrame:
    """
    Zgjedh PSU-të për një komunë:
    - garanton sa më shumë që të jetë e mundur përfaqësim të quadrant-eve
    - tenton të ketë të paktën një PSU ku ekziston çdo etni e kërkuar
    - shpërndan intervistat sipas rregullit të k/num_psu/leftover
    """

    df = df_psu_mun.copy()
    if df.empty or total_interviews <= 0:
        return pd.DataFrame()

    num_psu, leftover, extra_psu_size = compute_num_psu(total_interviews, k)
    if num_psu == 0:
        return pd.DataFrame()

    df["PopFilt"] = df.apply(
    lambda r: compute_filtered_pop_for_psu_row(
        r,
        age_min=min_age,
        age_max=max_age,
        gender_selected=gender_selected,
        eth_filter=eth_filter
    ),
    axis=1)

    ALL_ETH_COLS = [
        "Shqiptar", "Serb", "Boshnjak", "Turk", "Rom",
        "Ashkali", "Egjiptian", "Goran", "Të tjerë",
        "Preferoj të mos përgjigjem"
    ]

    OTHER_ETH_COLS = [e for e in ALL_ETH_COLS if e not in ["Shqiptar", "Serb"]]

    def compute_ethnic_pop_filtered(row):
        eth_total = sum(row.get(c, 0) for c in ALL_ETH_COLS)
        if eth_total <= 0:
            return pd.Series({
                "Shqiptar_pop": 0.0,
                "Serb_pop": 0.0,
                "Tjeter_pop": 0.0
            })

        shq = row.get("Shqiptar", 0) / eth_total
        ser = row.get("Serb", 0) / eth_total
        tjr = sum(row.get(c, 0) for c in OTHER_ETH_COLS) / eth_total

        return pd.Series({
            "Shqiptar_pop": row["PopFilt"] * shq,
            "Serb_pop":     row["PopFilt"] * ser,
            "Tjeter_pop":   row["PopFilt"] * tjr
        })

    # Remove any previous duplicate Tjeter_pop column
    if "Tjeter_pop" in df.columns:
        df = df.drop(columns=["Tjeter_pop"])

    eth_cols_df = df.apply(compute_ethnic_pop_filtered, axis=1)
    df = pd.concat([df, eth_cols_df], axis=1)

    # --------------------------
    # 1) Përfaqësimi i quadrant-eve
    # --------------------------
    df = df.sort_values("PopFilt", ascending=False)
    quads = df["Quadrant"].dropna().unique().tolist()

    selected_idx = []

    # a) nëse kemi mjaftueshëm PSU për të gjithë quadrant-et
    if num_psu >= len(quads):
        for q in quads:
            cand = df[df["Quadrant"] == q]
            if not cand.empty:
                selected_idx.append(cand.index[0])

        # plotëso numrin e PSU-ve me PSU-të më të mëdha të mbetura
        remaining_needed = num_psu - len(selected_idx)
        if remaining_needed > 0:
            remaining = df.drop(index=selected_idx)
            extra = remaining.head(remaining_needed).index.tolist()
            selected_idx.extend(extra)
    else:
        # b) nuk kemi mjaftueshëm PSU për të gjithë quadrant-et → greedy
        used_quads = set()
        for i, row in df.iterrows():
            q = row["Quadrant"]
            if q not in used_quads:
                selected_idx.append(i)
                used_quads.add(q)
                if len(selected_idx) == num_psu:
                    break
        # nëse akoma s'e kemi arritur numrin, plotëso me më të mëdhenjtë
        if len(selected_idx) < num_psu:
            remaining = df.drop(index=selected_idx)
            extra = remaining.head(num_psu - len(selected_idx)).index.tolist()
            selected_idx.extend(extra)

    selected_idx = list(dict.fromkeys(selected_idx))  # heq duplikate duke ruajtur rendin
    selected = df.loc[selected_idx].copy()

    # --------------------------
    # 2) Siguro prezencën e etnive të kërkuara
    # --------------------------
    def has_eth(selected_df, eth: str) -> bool:
        if eth == "Shqiptar":
            return (selected_df["Shqiptar_pop"] > 0).any()
        if eth == "Serb":
            return (selected_df["Serb_pop"] > 0).any()
        if eth == "Tjerë":
            return (selected_df["Tjeter_pop"] > 0).any()
        return False

    for eth in required_ethnicities:
        if has_eth(selected, eth):
            continue

        # gjej PSU jashtë të zgjedhurave që ka këtë etni
        if eth == "Shqiptar":
            cand_eth = df[(df["Shqiptar_pop"] > 0) & (~df.index.isin(selected.index))]
        elif eth == "Serb":
            cand_eth = df[(df["Serb_pop"] > 0) & (~df.index.isin(selected.index))]
        elif eth == "Tjerë":
            cand_eth = df[(df["Tjeter_pop"] > 0) & (~df.index.isin(selected.index))]
        else:
            continue

        if cand_eth.empty:
            # nuk ka PSU me këtë etni në këtë komunë
            continue

        # marrim kandidatin më të madh
        new_psu = cand_eth.iloc[0]

        # gjej një PSU për t'u zëvendësuar që nuk është i vetmi në quadrant-in e vet
        removed_idx = None
        for idx, row in selected.sort_values("PopFilt", ascending=False).iterrows():
            q = row["Quadrant"]
            # a ka PSU të tjera në të njëjtin quadrant në 'selected'?
            if (selected["Quadrant"] == q).sum() > 1:
                removed_idx = idx
                break

        if removed_idx is None:
            # nuk mund të bëjmë swap pa prishur quadrant-et → si fallback mund ta shtojmë
            # por për të mos prishur logjikën e numrit të PSU-ve, aktualisht e anashkalojmë
            continue

        # bëjmë zëvendësimin
        selected = selected.drop(index=removed_idx)
        selected = pd.concat([selected, new_psu.to_frame().T])

    # --------------------------
    # 3) Shpërndarja e intervistave te PSU-të
    # --------------------------
    selected = selected.sort_values("PopFilt", ascending=False).reset_index(drop=True)

    if extra_psu_size > 0:
        # p.sh. 46 anketa, k=8 → 6 PSU (5*8 + 6)
        base_sizes = [k] * (num_psu - 1) + [extra_psu_size]
    else:
        # p.sh. 42 anketa, k=8 → 5 PSU (5*8) + leftover=2 → shpërndajmë 2
        base_sizes = [k] * num_psu
        if leftover > 0:
            for i in range(min(leftover, len(base_sizes))):
                base_sizes[i] += 1

    selected["Intervista"] = base_sizes[: len(selected)]

    # shtojmë info etnish (num popullsie në atë PSU)
    selected["Shqiptar_pop"] = selected["Shqiptar_pop"].astype(float)
    selected["Serb_pop"] = selected["Serb_pop"].astype(float)
    selected["Tjeter_pop"] = selected["Tjeter_pop"].astype(float)

    selected["Komuna"] = komuna

    return selected[
        [
            "Komuna",
            "Fshati/Qyteti",
            "Vendbanimi",
            "Quadrant",
            "PopFilt",
            "Intervista",
            "Shqiptar_pop",
            "Serb_pop",
            "Tjeter_pop",
        ]
    ]


def compute_psu_table_for_all_municipalities(
    pivot: pd.DataFrame,
    df_psu: pd.DataFrame,
    k: int,
    eth_filter: list[str],
    settlement_filter: list[str]
) -> pd.DataFrame:
    """
    Gjeneron tabelën finale të PSU-ve për të gjitha komunat.
    - Urban = 1 rresht me total Urban intervista
    - Rural = përdor select_psus_for_municipality()
    """

    def extract_urban_interviews(pivot_row):
        urban_cols = [c for c in pivot_row.index if "Urban" in str(c)]
        return int(pivot_row[urban_cols].sum()) if urban_cols else 0

    # Gjej kolonat e etnisë në pivot
    eth_cols_map = {
        eth: [
            c for c in pivot.columns
            if c != "Total" and str(c).startswith(eth)
        ]
        for eth in ["Shqiptar", "Serb", "Tjerë"]
    }

    all_rows = []

    for kom in pivot.index:
        if kom == "Total":
            continue

        # pikënisja
        pivot_row = pivot.loc[kom]
        total_interviews = int(pivot_row["Total"])
        if total_interviews <= 0:
            continue

        # Llogarit Urban dhe Rural
        urban_int = extract_urban_interviews(pivot_row)
        rural_int = total_interviews - urban_int

        df_mun = df_psu[df_psu["Komuna"] == kom].copy()
        if df_mun.empty:
            continue

        # Filtrim sipas vendbanimit të zgjedhur (nëse e përdor në app)
        if settlement_filter:
            df_mun = df_mun[df_mun["Vendbanimi"].isin(settlement_filter)]

        # ===========================
        # 1) URBAN PSU (një rresht)
        # ===========================
        # ===========================
        # 1) URBAN PSU (always single)
        # ===========================
        if urban_int > 0:
            df_mun_urban = df_mun[df_mun["Vendbanimi"] == "Urban"]

            if not df_mun_urban.empty:
                # There is always exactly 1 Urban row per municipality
                best_urban = df_mun_urban.iloc[0].copy()

                # Compute PopFilt for Urban row
                best_urban["PopFilt"] = compute_filtered_pop_for_psu_row(
                    best_urban,
                    age_min=min_age,
                    age_max=max_age,
                    gender_selected=gender_selected,
                    eth_filter=eth_filter
                )


                row_urban = pd.DataFrame([{
                    "Komuna": kom,
                    "Fshati/Qyteti": best_urban["Fshati/Qyteti"],
                    "Vendbanimi": "Urban",
                    "Quadrant": "-",
                    "PopFilt": best_urban["PopFilt"],
                    "Intervista": urban_int,
                    "Shqiptar_pop": best_urban.get("Shqiptar_pop", 0),
                    "Serb_pop": best_urban.get("Serb_pop", 0),
                    "Tjeter_pop": best_urban.get("Tjeter_pop", 0)
                }])

                all_rows.append(row_urban)


        # ===========================
        # 2) RURAL PSU (përdor logjikën e tanishme)
        # ===========================

        # Gjej cilat etni kanë mostra > 0 në këtë komunë
        required_eth = []
        for eth, cols in eth_cols_map.items():
            if eth not in eth_filter:
                continue
            if not cols:
                continue
            if int(pivot.loc[kom, cols].sum()) > 0:
                required_eth.append(eth)

        if rural_int > 0:
            df_mun_rural = df_mun[df_mun["Vendbanimi"] == "Rural"]

            psu_rural = select_psus_for_municipality(
                komuna=kom,
                total_interviews=rural_int,
                df_psu_mun=df_mun_rural,
                k=k,
                required_ethnicities=required_eth
            )

            if not psu_rural.empty:
                all_rows.append(psu_rural)

    # ===========================
    # 3) Bashkimi final
    # ===========================
    if not all_rows:
        return pd.DataFrame()

    final_psu = pd.concat(all_rows, ignore_index=True)
    return final_psu

def extract_urban_interviews(pivot_row):
    urban_cols = [c for c in pivot_row.index if "Urban" in str(c)]
    return int(pivot_row[urban_cols].sum()) if urban_cols else 0

def compute_filtered_pop_for_psu_row(
    psu_row: pd.Series,
    age_min: int,
    age_max: int | None,
    gender_selected: list[str],
    eth_filter: list[str]
) -> float:
    """
    Compute population for one PSU row using demographic filters applied directly to df_psu.
    """

    total_pop = 0

    # 1. Handle age max
    if age_max is None:
        age_max = 120  # big number so it includes everything

    # 2. Identify all age group columns (e.g. "0-4", "5-9", ...)
    age_cols = []
    for col in psu_row.index:
        name = str(col).strip()
        if re.fullmatch(r"\d+\-\d+", name) or re.fullmatch(r"\d+\+", name):
            age_cols.append(col)

    def group_range(col_name):
        if "+" in col_name:
            base = int(col_name.replace("+", "").strip())
            return (base, 200)
        else:
            a, b = col_name.split("-")
            return (int(a), int(b))

    # 3. Loop through age groups
    for col in age_cols:
        lo, hi = group_range(col)
        group_pop = psu_row[col]

        # skip if no population
        if group_pop <= 0:
            continue

        # calculate overlap between group [lo,hi] and filter [age_min,age_max]
        overlap_lo = max(lo, age_min)
        overlap_hi = min(hi, age_max)

        if overlap_lo > overlap_hi:
            # no overlap
            continue

        # fraction of the group included
        group_size = hi - lo + 1
        overlap_size = overlap_hi - overlap_lo + 1
        fraction = overlap_size / group_size

        total_pop += group_pop * fraction

    # 4. Gender filter
    if gender_selected == ["Meshkuj"]:
        gender_factor = psu_row["Meshkuj"] / (psu_row["Meshkuj"] + psu_row["Femra"])
        total_pop *= gender_factor

    elif gender_selected == ["Femra"]:
        gender_factor = psu_row["Femra"] / (psu_row["Meshkuj"] + psu_row["Femra"])
        total_pop *= gender_factor

    # If both genders selected → keep full total_pop

    # 5. Ethnicity filter
    eth_pop = 0
    for eth in eth_filter:
        if eth in psu_row:
            eth_pop += psu_row[eth]

    # denominator = total population in PSU (sum of all ethnic groups)
    all_ethnic_cols = [
        "Shqiptar", "Serb", "Boshnjak", "Turk", "Rom",
        "Ashkali", "Egjiptian", "Goran", "Të tjerë",
        "Preferoj të mos përgjigjem"
    ]

    eth_total = sum(psu_row.get(e, 0) for e in all_ethnic_cols)

    if eth_total > 0:
        total_pop *= (eth_pop / eth_total)

    return total_pop

# Load data
try:
    df_eth = load_ethnicity_settlement_data("excel-files/ASK-2024-Komuna-Etnia-Vendbanimi.xlsx")
    df_ga, age_cols = load_gender_age_data("excel-files/ASK-2024-Komuna-Gjinia-Mosha.xlsx")
except Exception as e:
    st.error(f"Gabim gjatë leximit të fajllave: {e}")
    st.stop()

region_map = get_region_mapping()

try:
    df_psu = load_psu_data("excel-files/ASK-2024-Komuna-Vendbanim-Fshat+Qytet.xlsx")
except Exception as e:
    st.error(f"Gabim gjatë leximit të fajllit të PSU-ve: {e}")
    st.stop()

# =========================
# UI: SIDEBAR
# =========================

st.title("Dizajnimi i Mostrës Nacionale")

st.sidebar.header("Parametrat kryesorë")

# Total sample size
n_total = st.sidebar.number_input(
    "Numri total i mostrës (N)",
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
st.sidebar.subheader("Mbledhja e të dhënave")

data_collection_method = st.sidebar.selectbox(
    "Metoda e mbledhjes së të dhënave",
    options=["CAPI", "CATI", "CAWI"],
    index=0
)

if data_collection_method=="CAPI":
    interviews_per_psu = st.sidebar.slider(
        "Numri i intervistave për PSU",
        min_value=6,
        max_value=12,
        value=8,
        step=1
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
    "Gjinia që përfshihet",
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

        # ============================
        # 1) SPECIAL CASE: MOSHA
        # ============================
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
                f"Numri i anketave për {min_over_age}–{max_over_age}",
                min_value=1, value=50, step=10,
                key=f"n_{var}"
            )

            oversample_inputs[var] = [{
                "min_age": min_over_age,
                "max_age": max_over_age,
                "n": oversample_n
            }]
            continue

        # ============================
        # 2) Merr opsionet e vlefshme
        # ============================
        df_tmp = df_eth[
            (df_eth["Etnia"].isin(eth_filter)) &
            (df_eth["Vendbanimi"].isin(settlement_filter)) &
            (df_eth["Komuna"].isin(komuna_filter))
        ]

        if var == "Komuna":
            valid_kom = df_tmp.groupby("Komuna")["Pop_base"].sum()
            valid_kom = valid_kom[valid_kom > 0].index.tolist()
            options = sorted(valid_kom)
            allow_multiple = True

        elif var == "Etnia":
            options = sorted(eth_filter)
            allow_multiple = True

        elif var == "Regjion":
            valid_kom = df_tmp.groupby("Komuna")["Pop_base"].sum()
            valid_kom = valid_kom[valid_kom > 0].index.tolist()
            valid_reg = {region_map[k] for k in valid_kom if k in region_map}
            options = sorted(valid_reg)
            allow_multiple = False

        elif var == "Vendbanimi":
            options = sorted(settlement_filter)
            allow_multiple = False

        elif var == "Gjinia":
            options = sorted(gender_selected)
            allow_multiple = False

        else:
            options = []
            allow_multiple = False

        # ============================
        # 3) UI: multiselect vetëm për Komuna/Etnia
        # ============================
        if allow_multiple:
            selected_values = st.sidebar.multiselect(
                f"Zgjidh {var} për oversample (Mund të zgjidhni më shumë se një)",
                options=options,
                key=f"multi_{var}"
            )

            entry_list = []
            for v in selected_values:
                q = st.sidebar.number_input(
                    f"Kuota për {var} = {v}",
                    min_value=1, value=50, step=10,
                    key=f"quota_{var}_{v}"
                )
                entry_list.append({"value": v, "n": q})

            oversample_inputs[var] = entry_list

        else:
            val = st.sidebar.selectbox(
                f"Njësia nga {var} që do të jetë oversample",
                options=options,
                key=f"val_{var}"
            )
            q = st.sidebar.number_input(
                f"Kuota për {var} = {val}",
                min_value=1, value=50, step=10,
                key=f"quota_{var}_{val}"
            )

            oversample_inputs[var] = [{"value": val, "n": q}]


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
    for var, entry_list in oversample_inputs.items():
        for entry in entry_list:
            precomputed_masks[var] = mask_for_oversample(grouped, var, entry)

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

    all_os = []

    for var, entries in oversample_inputs.items():
        for entry in entries:

            mask = mask_for_oversample(grouped, var, entry)

            # MOSHA ka strukturë ndryshe → nuk ka "value"
            if var == "Mosha":
                all_os.append({
                    "var": var,
                    "value": f"{entry['min_age']}-{entry['max_age']}",
                    "n": entry["n"],
                    "mask": mask
                })
            else:
                # Komuna, Etnia, Regjioni, Vendbanimi, Gjinia
                all_os.append({
                    "var": var,
                    "value": entry["value"],
                    "n": entry["n"],
                    "mask": mask
                })

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
        varA, entry_list = oversample_items[0]

        # fillo me zero
        grouped["n_alloc"] = 0

        # për çdo kuotë të atij variabli
        used_total = 0
        for entry in entry_list:
            nA = int(entry["n"])
            maskA = mask_for_oversample(grouped, varA, entry)

            alloc_A = alloc_to_mask(maskA, nA)
            grouped["n_alloc"] += alloc_A
            used_total += int(alloc_A.sum())

        # pjesa e mbetur shkon tek stratat tjera
        remaining = n_total - used_total
        if remaining < 0:
            remaining = 0

        mask_rest = (grouped["n_alloc"] == 0)
        alloc_rest = alloc_to_mask(mask_rest, remaining)

        grouped["n_alloc"] += alloc_rest

    # 2) Dy variabla oversample (me shumë vlera për njërin variabël)
    else:
        # ndërto listën e plotë që i ke më lart
        # all_os = [ {var, value, n, mask}, ... ]

        # Renditi sipas kuotës zbritëse
        all_os_sorted = sorted(all_os, key=lambda x: x["n"], reverse=True)

        # OS_B = variabli me kuotën më të lartë (p.sh. Rural=800)
        osB = all_os_sorted[0]

        # OS_A = të gjithë tjerët (p.sh. Peja=300, Prishtina=500, Gjakova=200)
        osA_list = all_os_sorted[1:]

        # 1) shpërndaj OS_B
        alloc_B = alloc_to_mask(osB["mask"], osB["n"])

        grouped["n_alloc"] = alloc_B

        # 2) pastaj secilin OS_A një nga një
        for osA in osA_list:

            # llogarit overlapp
            overlap_mask = osA["mask"] & osB["mask"]
            overlap_from_B = int(alloc_B[overlap_mask].sum())

            remaining_A = max(osA["n"] - overlap_from_B, 0)

            alloc_A = alloc_to_mask(osA["mask"] & ~osB["mask"], remaining_A)

            grouped["n_alloc"] += alloc_A

        # 3) pjesa e mbetur shkon jashtë OS-ve
        used = int(grouped["n_alloc"].sum())
        remaining = max(n_total - used, 0)

        mask_rest = ~(sum(os["mask"] for os in all_os) > 0)
        alloc_rest = alloc_to_mask(mask_rest, remaining)

        grouped["n_alloc"] += alloc_rest

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

        os_gender = oversample_inputs["Gjinia"][0]["value"]
        os_n = int(oversample_inputs["Gjinia"][0]["n"])

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

        params_age = oversample_inputs["Mosha"][0]
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

    pivot_old = pivot.copy()
    
    pivot = fix_minimum_allocations(
            pivot=pivot,
            df_eth= df_eth,
            region_map=region_map,
            min_total=3,   # minimum anketa per komunë
            min_eth=3      # minimum per vendbanim (Urban/Rural)
        )
    
    # Safety: ensure global total matches n_total
    global_total = int(pivot.loc["Total", "Total"])

    st.subheader("Tabela e ndarjes së mostrës")

    # Përgatit tekstin për grupmoshën
    if max_age is None:
        age_text = f"Grupmosha: **{min_age}+**"
    else:
        age_text = f"Grupmosha: **{min_age}–{max_age}**"

    # Përgatit tekstin për gjininë
    if len(gender_selected) == 1:
        gender_text = f"Gjinia: **{gender_selected[0]}**"
    else:
        gender_text = ""

    if len(settlement_filter) == 1:
        settlement_text = f"Vendbanimi: **{settlement_filter[0]}**"
    else:
        settlement_text = ""


    # Linja kryesore
    caption_main = (
        f"Ndarja kryesore: **{primary_level}** | "
        f"Nën-ndarja: **{', '.join(sub_options) if sub_options else 'Asnjë'}** | "
        f"Totali i mostrës: **{n_total}** | "
        f"Totali i alokuar: **{global_total}**"
    )

    # Linja shtesë për filtrat demografikë
    caption_extra = " | ".join(filter(None, [age_text, gender_text, settlement_text]))

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
    create_download_link2(
        file_bytes=strata_excel,
        filename="mostra_strata.xlsx",
        label="Shkarko Strata"
    )

    # 📘 Old Pivot table (Excel)
    strata_excel = df_to_excel_bytes(pivot_old, sheet_name="Shpërndarja fillestare")
    create_download_link2(
        file_bytes=strata_excel,
        filename="shpërndarja_fillestare.xlsx",
        label="Shkarko Shpërndarjen Fillestare"
    )

        # =====================================================
    # PSU-të vetëm nëse metoda është CAPI dhe niveli kryesor është Komunë
    # =====================================================
    if data_collection_method == "CAPI":
        if primary_level != "Komunë":
            st.info("Llogaritja e PSU-ve është e implementuar vetëm kur ndarja kryesore është sipas **Komunës**.")
        else:
            st.subheader("PSU-të e përzgjedhura")

            with st.spinner("Duke llogaritur PSU-të..."):
                psu_table = compute_psu_table_for_all_municipalities(
                    pivot=pivot,
                    df_psu=df_psu,
                    k=interviews_per_psu,
                    eth_filter=eth_filter,
                    settlement_filter=settlement_filter,
                )


            if psu_table.empty:
                st.warning("Nuk u gjeneruan PSU. Kontrollo filtrat, fajllin e PSU-ve dhe shpërndarjen e mostrës.")
            else:
                st.caption(
                    f"PSU-të janë llogaritur me **{interviews_per_psu} intervista** për PSU sipas rregullit të përcaktuar."
                )
                st.dataframe(psu_table, use_container_width=True)

                psu_excel = df_to_excel_bytes(psu_table, sheet_name="PSU")
                create_download_link2(
                    file_bytes=psu_excel,
                    filename="psu_capi_tegjitha_komunat.xlsx",
                    label="Shkarko PSU-të"
                )

else:
    st.info("Cakto parametrat kryesorë dhe kliko **'Gjenero shpërndarjen e mostrës'** për të dizajnuar mostrën.")
