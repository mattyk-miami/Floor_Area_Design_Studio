import io
import re
import copy
import math
import json
from pathlib import Path
import pandas as pd
import streamlit as st
import xlsxwriter  # Explicit dependency for Excel export

st.set_page_config(page_title="Floor Area Design Studio", page_icon="🏢", layout="wide")

APP_VERSION = "2.3.0"
SCENARIO_FILE = Path(__file__).with_name("floor_area_scenarios.json")


# ---------- Stormy Morning theme ----------
STORMY = {
    "slate": "#6A89A7",
    "mist": "#BDDDFC",
    "sky": "#88BDF2",
    "charcoal": "#384959",
    "page": "#F7FAFD",
}

st.markdown(
    f"""
    <style>
    :root {{
        --stormy-slate: {STORMY['slate']};
        --stormy-mist: {STORMY['mist']};
        --stormy-sky: {STORMY['sky']};
        --stormy-charcoal: {STORMY['charcoal']};
        --stormy-page: {STORMY['page']};
    }}

    /* Main page */
    [data-testid="stAppViewContainer"] {{
        background: linear-gradient(180deg, #FFFFFF 0%, var(--stormy-page) 100%);
        color: var(--stormy-charcoal);
    }}
    [data-testid="stHeader"] {{
        background: rgba(255,255,255,0.88);
        backdrop-filter: blur(8px);
    }}
    .block-container {{
        padding-top: 2.0rem;
        padding-bottom: 3rem;
        max-width: 1600px;
    }}

    /* Typography */
    h1, h2, h3, h4 {{ color: var(--stormy-charcoal) !important; }}
    h1 {{ letter-spacing: -0.035em; }}
    p, label {{ color: var(--stormy-charcoal); }}
    [data-testid="stCaptionContainer"] p {{ color: var(--stormy-slate) !important; }}

    /* Sidebar */
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, var(--stormy-charcoal) 0%, #49647A 62%, var(--stormy-slate) 100%);
        border-right: 1px solid rgba(255,255,255,0.12);
    }}
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{
        color: #F7FBFF !important;
    }}
    [data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.22); }}

    /* Inputs */
    [data-baseweb="input"] > div,
    [data-baseweb="select"] > div,
    [data-baseweb="base-input"] {{
        border-color: var(--stormy-slate) !important;
        border-radius: 8px !important;
    }}
    [data-testid="stSidebar"] [data-baseweb="input"] > div,
    [data-testid="stSidebar"] [data-baseweb="select"] > div,
    [data-testid="stSidebar"] [data-baseweb="base-input"] {{
        background: rgba(255,255,255,0.96) !important;
    }}

    /* Buttons */
    .stButton > button,
    .stDownloadButton > button {{
        border-radius: 8px;
        border: 1px solid var(--stormy-slate);
        background: var(--stormy-mist);
        color: var(--stormy-charcoal);
        font-weight: 650;
        transition: all 0.15s ease;
    }}
    .stButton > button:hover,
    .stDownloadButton > button:hover {{
        background: var(--stormy-sky);
        border-color: var(--stormy-charcoal);
        color: var(--stormy-charcoal);
        transform: translateY(-1px);
    }}
    [data-testid="stSidebar"] .stButton > button {{
        background: rgba(189,221,252,0.95);
        color: var(--stormy-charcoal);
        border-color: rgba(255,255,255,0.32);
    }}

    /* Metric cards */
    [data-testid="stMetric"] {{
        background: linear-gradient(145deg, rgba(189,221,252,0.48), rgba(136,189,242,0.18));
        border: 1px solid rgba(106,137,167,0.38);
        border-radius: 12px;
        padding: 14px 16px;
        box-shadow: 0 3px 12px rgba(56,73,89,0.06);
        min-height: 106px;
    }}
    [data-testid="stMetricLabel"] {{
        color: var(--stormy-slate) !important;
        overflow: visible !important;
        white-space: normal !important;
        text-overflow: clip !important;
    }}
    [data-testid="stMetricLabel"] > div,
    [data-testid="stMetricLabel"] p {{
        overflow: visible !important;
        white-space: normal !important;
        text-overflow: clip !important;
        line-height: 1.15 !important;
    }}
    [data-testid="stMetricValue"] {{
        color: var(--stormy-charcoal) !important;
        overflow: visible !important;
        width: 100% !important;
    }}
    [data-testid="stMetricValue"] > div,
    [data-testid="stMetricValue"] p {{
        color: var(--stormy-charcoal) !important;
        overflow: visible !important;
        text-overflow: clip !important;
        white-space: nowrap !important;
        max-width: none !important;
        width: auto !important;
        font-size: clamp(1.45rem, 2vw, 1.95rem) !important;
        line-height: 1.15 !important;
    }}

    /* Tabs */
    [data-baseweb="tab-list"] {{
        gap: 8px;
        border-bottom: 1px solid var(--stormy-mist);
    }}
    [data-baseweb="tab"] {{
        background: rgba(189,221,252,0.30);
        border-radius: 8px 8px 0 0;
        padding-left: 16px;
        padding-right: 16px;
        color: var(--stormy-charcoal);
    }}
    [aria-selected="true"][data-baseweb="tab"] {{
        background: var(--stormy-mist);
        color: var(--stormy-charcoal) !important;
        font-weight: 700;
    }}

    /* Tables/editors/expanders */
    [data-testid="stDataFrame"],
    [data-testid="stDataEditor"],
    [data-testid="stExpander"] {{
        border: 1px solid rgba(106,137,167,0.30);
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 10px rgba(56,73,89,0.04);
    }}
    [data-testid="stExpander"] details {{ background: rgba(255,255,255,0.72); }}

    /* Subtle section separators */
    hr {{ border-color: rgba(106,137,167,0.26); }}

    /* Keep semantic compliance alerts recognizable, but soften the container edges */
    [data-testid="stAlert"] {{ border-radius: 10px; }}
    </style>
    """,
    unsafe_allow_html=True,
)

UNIT_TYPES = {
    "S-A1": 508,
    "1_Bd-B1": 618,
    "1_Bd-B2": 617,
    "1_Bd-B3": 751,
    "1_Bd-B4": 706,
    "2_Bd-C1": 995,
    "2_Bd-C2": 1024,
    "2_Bd-C3": 921,
    "2_Bd-C4": 1026,
    "2_Bd-C5": 1043,
    "2_Bd-C6": 1088,
    "3_Bd-D1": 1130,
}
UNIT_COLS = list(UNIT_TYPES)

PROJECT_DEFAULTS = {
    "building_name_v2": "Floor Area Study",
    "lot_area_v2": 43050.0,
    "allowable_far_v2": 4.75,
    "retail_sf_v2": 4677.0,
    "p_offstreet_v2": 0,
    "loading_provided_v2": 0,
}


def initial_floors():
    rows = []
    for level in range(13, 5, -1):
        rows.append({
            "Floor": f"Residential Level {level}",
            "Type": "Residential",
            "Garage / Pool Deck": 0,
            "Enclosed": 19860,
            "Balconies": 0,
            "Total GSF": 19860,
            "LSF": 17424,
            "GFA": 19860,
            "Parking Spaces": 0,
            "S-A1": 2, "1_Bd-B1": 8, "1_Bd-B2": 2, "1_Bd-B3": 0, "1_Bd-B4": 0,
            "2_Bd-C1": 2, "2_Bd-C2": 2, "2_Bd-C3": 4, "2_Bd-C4": 0, "2_Bd-C5": 0,
            "2_Bd-C6": 0, "3_Bd-D1": 2,
        })
    rows.append({
        "Floor": "Residential/Amenities Level 5", "Type": "Residential / Amenity",
        "Garage / Pool Deck": 17508, "Enclosed": 19396, "Balconies": 0,
        "Total GSF": 33937, "LSF": 13680, "GFA": 19396,
        "Parking Spaces": 0,
        "S-A1": 2, "1_Bd-B1": 6, "1_Bd-B2": 2, "1_Bd-B3": 0, "1_Bd-B4": 0,
        "2_Bd-C1": 2, "2_Bd-C2": 2, "2_Bd-C3": 4, "2_Bd-C4": 0, "2_Bd-C5": 0,
        "2_Bd-C6": 0, "3_Bd-D1": 0,
    })
    parking_by_level = {4: 106, 3: 100, 2: 72}
    for level in (4, 3, 2):
        rows.append({
            "Floor": f"Parking Level {level}", "Type": "Parking",
            "Garage / Pool Deck": 32812, "Enclosed": 0, "Balconies": 0,
            "Total GSF": 33937, "LSF": 0, "GFA": 601,
            "Parking Spaces": parking_by_level[level],
            **{u: 0 for u in UNIT_COLS},
        })
    rows.append({
        "Floor": "Ground Level", "Type": "Ground / Podium",
        "Garage / Pool Deck": 7046, "Enclosed": 21590, "Balconies": 0,
        "Total GSF": 32485, "LSF": 4668, "GFA": 21590,
        "Parking Spaces": 8,
        "S-A1": 1, "1_Bd-B1": 0, "1_Bd-B2": 0, "1_Bd-B3": 0, "1_Bd-B4": 3,
        "2_Bd-C1": 0, "2_Bd-C2": 0, "2_Bd-C3": 0, "2_Bd-C4": 0, "2_Bd-C5": 2,
        "2_Bd-C6": 0, "3_Bd-D1": 0,
    })
    return pd.DataFrame(rows)


def default_unit_sizes_df():
    def category(unit):
        if unit.startswith("S-"):
            return "Studio"
        if unit.startswith("1_"):
            return "1 Bedroom"
        if unit.startswith("2_"):
            return "2 Bedroom"
        return "3 Bedroom"

    return pd.DataFrame({
        "Unit Type": UNIT_COLS,
        "Category": [category(u) for u in UNIT_COLS],
        "Unit Size (SF)": [UNIT_TYPES[u] for u in UNIT_COLS],
    })


def unit_size_dict(df):
    temp = df.copy()
    temp["Unit Size (SF)"] = pd.to_numeric(temp["Unit Size (SF)"], errors="coerce").fillna(0)
    return dict(zip(temp["Unit Type"], temp["Unit Size (SF)"].astype(float)))


def floor_sort_value(name):
    s = str(name)
    if "Ground" in s:
        return -1
    m = re.search(r"Level\s+(\d+)", s, flags=re.I)
    return int(m.group(1)) if m else 0


def sort_building(df):
    return (
        df.assign(_sort=df["Floor"].map(floor_sort_value))
        .sort_values(["_sort", "Floor"], ascending=[False, True])
        .drop(columns="_sort")
        .reset_index(drop=True)
    )


def recalc(df, unit_sizes):
    out = df.copy()
    numeric = ["Garage / Pool Deck", "Enclosed", "Balconies", "Total GSF", "LSF", "GFA", "Parking Spaces"] + UNIT_COLS
    for c in numeric:
        if c not in out.columns:
            out[c] = 0
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0)
    out[UNIT_COLS] = out[UNIT_COLS].clip(lower=0).round().astype(int)
    out["Parking Spaces"] = out["Parking Spaces"].clip(lower=0).round().astype(int)
    out["Units"] = out[UNIT_COLS].sum(axis=1).astype(int)
    out["Unit SF"] = sum(out[u] * float(unit_sizes.get(u, 0)) for u in UNIT_COLS)
    return out


def make_blank_floor(existing_df):
    levels = [floor_sort_value(x) for x in existing_df["Floor"]]
    next_level = max([x for x in levels if x >= 1], default=0) + 1
    row = {
        "Floor": f"Residential Level {next_level}",
        "Type": "Residential",
        "Garage / Pool Deck": 0,
        "Enclosed": 0,
        "Balconies": 0,
        "Total GSF": 0,
        "LSF": 0,
        "GFA": 0,
        "Parking Spaces": 0,
        **{u: 0 for u in UNIT_COLS},
    }
    return row


def unique_duplicate_name(source_name, existing_names):
    match = re.search(r"(.*Level\s+)(\d+)(.*)", source_name, flags=re.I)
    if match:
        candidate = f"{match.group(1)}{int(match.group(2)) + 1}{match.group(3)}"
        if candidate not in existing_names:
            return candidate
    base = f"{source_name} Copy"
    candidate = base
    n = 2
    while candidate in existing_names:
        candidate = f"{base} {n}"
        n += 1
    return candidate






def duplicate_floor_in_sequence(df, source_name):
    """Duplicate a numeric floor directly above itself and keep levels continuous.

    If Parking Level 4 is duplicated, the copy becomes Parking Level 5 and
    every existing numeric floor at Level 5 or higher shifts up one level.
    This keeps parking below amenity/residential floors instead of sending the
    duplicated parking floor to an unrelated position in the stack.
    """
    out = df.copy()
    match = re.search(r"Level\s+(\d+)", str(source_name), flags=re.I)

    # Non-numbered floors (normally Ground) retain the older copy-name behavior.
    if not match:
        source = out.loc[out["Floor"] == source_name].iloc[0].copy()
        new_name = unique_duplicate_name(str(source["Floor"]), set(out["Floor"]))
        source["Floor"] = new_name
        out = pd.concat([out, pd.DataFrame([source])], ignore_index=True)
        return sort_building(out), new_name

    source_level = int(match.group(1))
    insert_level = source_level + 1

    def shift_name(name):
        text = str(name)
        m = re.search(r"Level\s+(\d+)", text, flags=re.I)
        if not m:
            return text
        level = int(m.group(1))
        if level >= insert_level:
            level += 1
            return text[:m.start(1)] + str(level) + text[m.end(1):]
        return text

    # Shift everything above the selected floor first.
    out["Floor"] = out["Floor"].map(shift_name)

    # The source row itself did not move, so copy all its data (including
    # parking spaces and unit mix) into the newly opened level immediately above.
    source = out.loc[out["Floor"] == source_name].iloc[0].copy()
    source_text = str(source["Floor"])
    m = re.search(r"Level\s+(\d+)", source_text, flags=re.I)
    new_name = source_text[:m.start(1)] + str(insert_level) + source_text[m.end(1):]
    source["Floor"] = new_name

    out = pd.concat([out, pd.DataFrame([source])], ignore_index=True)
    return sort_building(out), new_name


def renumber_floors(df):
    """Compact numeric floor levels after a deletion while preserving floor use labels.

    Example: Parking Level 3, Parking Level 4, Amenity Level 5 becomes
    Parking Level 3, Amenity Level 4 if Level 4 was deleted. Ground is
    intentionally left unchanged.
    """
    out = df.copy()

    # Collect the distinct numeric levels that remain in the building.
    levels = []
    for name in out["Floor"].astype(str):
        m = re.search(r"Level\s+(\d+)", name, flags=re.I)
        if m:
            levels.append(int(m.group(1)))

    levels = sorted(set(levels))
    if not levels:
        return sort_building(out)

    # Keep the building's current lowest numbered level as the anchor, then
    # close every gap above it. In the source model that anchor is Level 2.
    first_level = levels[0]
    level_map = {old: first_level + i for i, old in enumerate(levels)}

    def rename(name):
        s = str(name)
        m = re.search(r"Level\s+(\d+)", s, flags=re.I)
        if not m:
            return s
        old = int(m.group(1))
        new = level_map.get(old, old)
        return s[:m.start(1)] + str(new) + s[m.end(1):]

    out["Floor"] = out["Floor"].map(rename)
    return sort_building(out)


def normalize_floors(df):
    out = df.copy()
    if "Parking Spaces" not in out.columns:
        out["Parking Spaces"] = 0
    for col in ["Parking Spaces"] + UNIT_COLS:
        if col not in out.columns:
            out[col] = 0
    return sort_building(out)


def serialize_snapshot(snapshot):
    return {
        "floors": snapshot["floors"].to_dict(orient="records"),
        "unit_sizes_df": snapshot["unit_sizes_df"].to_dict(orient="records"),
        "project": snapshot["project"],
    }


def deserialize_snapshot(data):
    project = dict(data.get("project", {}))
    # Backward compatibility with scenarios created before the Site -> Lot rename.
    if "lot_area_v2" not in project and "site_area_v2" in project:
        project["lot_area_v2"] = project.pop("site_area_v2")
    return {
        "floors": normalize_floors(pd.DataFrame(data.get("floors", []))),
        "unit_sizes_df": pd.DataFrame(data.get("unit_sizes_df", [])),
        "project": project,
    }


def load_persistent_scenarios():
    if not SCENARIO_FILE.exists():
        return {}
    try:
        raw = json.loads(SCENARIO_FILE.read_text(encoding="utf-8"))
        return {name: deserialize_snapshot(snap) for name, snap in raw.items()}
    except Exception:
        return {}


def persist_scenarios():
    payload = {name: serialize_snapshot(snap) for name, snap in st.session_state.scenarios.items()}
    SCENARIO_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def bump_editor_keys():
    st.session_state.floor_editor_rev += 1
    st.session_state.mix_editor_rev += 1
    st.session_state.size_editor_rev += 1


def snapshot_current():
    return {
        "floors": st.session_state.floors.copy(deep=True),
        "unit_sizes_df": st.session_state.unit_sizes_df.copy(deep=True),
        "project": {k: st.session_state.get(k, v) for k, v in PROJECT_DEFAULTS.items()},
    }


def queue_snapshot_load(name):
    # Apply saved widget values at the start of the next run, before those
    # widgets are instantiated. This avoids Streamlit session-state errors.
    st.session_state.pending_snapshot_name = name


def render_building_stack(calc):
    if calc.empty:
        st.info("No floors to display.")
        return
    max_gsf = max(float(calc["Total GSF"].max()), 1.0)
    type_fill = {
        "Residential": (STORMY["sky"], STORMY["charcoal"]),
        "Residential / Amenity": (STORMY["mist"], STORMY["charcoal"]),
        "Parking": (STORMY["charcoal"], "#FFFFFF"),
        "Ground / Podium": (STORMY["slate"], "#FFFFFF"),
        "Amenity": (STORMY["mist"], STORMY["charcoal"]),
        "Other": ("#DCE7F0", STORMY["charcoal"]),
    }
    html = ["<div style='max-width:820px;margin:0 auto;padding:8px 0;'>"]
    for _, row in calc.iterrows():
        width = max(48, int(96 * float(row["Total GSF"]) / max_gsf))
        fill, text_color = type_fill.get(str(row["Type"]), ("#DCE7F0", STORMY["charcoal"]))
        html.append(
            f"<div title='{row['Floor']}' style='width:{width}%;margin:4px auto;"
            f"background:{fill};color:{text_color};border:1px solid {STORMY['slate']};border-radius:7px;padding:9px 12px;"
            f"box-shadow:0 2px 6px rgba(56,73,89,0.10);display:flex;justify-content:space-between;gap:16px;font-size:0.92rem;'>"
            f"<strong>{row['Floor']}</strong>"
            f"<span>{row['Units']:.0f} units &nbsp;•&nbsp; {int(row['Parking Spaces'])} parking &nbsp;•&nbsp; {row['Total GSF']:,.0f} GSF</span></div>"
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


# ---------- Session initialization ----------
if "floors" not in st.session_state:
    st.session_state.floors = normalize_floors(initial_floors())
if "unit_sizes_df" not in st.session_state:
    st.session_state.unit_sizes_df = default_unit_sizes_df()
if "scenarios" not in st.session_state:
    st.session_state.scenarios = load_persistent_scenarios()
if "floor_editor_rev" not in st.session_state:
    st.session_state.floor_editor_rev = 0
if "mix_editor_rev" not in st.session_state:
    st.session_state.mix_editor_rev = 0
if "size_editor_rev" not in st.session_state:
    st.session_state.size_editor_rev = 0
if "selected_floor" not in st.session_state:
    st.session_state.selected_floor = st.session_state.floors.iloc[0]["Floor"]
if "pending_selected_floor" not in st.session_state:
    st.session_state.pending_selected_floor = None
if "pending_snapshot_name" not in st.session_state:
    st.session_state.pending_snapshot_name = None
for key, default in PROJECT_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = default

# Apply any queued scenario before project widgets are created.
pending_snapshot = st.session_state.get("pending_snapshot_name")
if pending_snapshot and pending_snapshot in st.session_state.scenarios:
    snap = st.session_state.scenarios[pending_snapshot]
    st.session_state.floors = snap["floors"].copy(deep=True)
    st.session_state.unit_sizes_df = snap["unit_sizes_df"].copy(deep=True)
    for k, v in snap["project"].items():
        st.session_state[k] = v
    st.session_state.pending_selected_floor = st.session_state.floors.iloc[0]["Floor"]
    st.session_state.pending_snapshot_name = None
    bump_editor_keys()


# ---------- Header ----------
st.title("Floor Area Design Studio")
building_name = st.text_input("Building Name", key="building_name_v2", placeholder="Enter building or project name")
if building_name.strip():
    st.markdown(f"## {building_name.strip()}")
st.caption(f"App version {APP_VERSION} • Building-first workflow • Top floor to Ground")

# ---------- Sidebar ----------
with st.sidebar:
    st.header("Project Inputs")
    lot_area = st.number_input("Lot area (SF)", min_value=0.0, step=50.0, key="lot_area_v2", help="Used for FAR = Total GFA ÷ Lot Area and Allowable GFA = Lot Area × Allowable FAR.")
    allowable_far = st.number_input("Allowable FAR", min_value=0.0, step=0.05, key="allowable_far_v2")
    retail_sf = st.number_input("Retail area (SF)", min_value=0.0, step=50.0, key="retail_sf_v2")

    st.divider()
    st.subheader("Parking Provided")
    p_offstreet = st.number_input("Off-street / site parking", min_value=0, step=1, key="p_offstreet_v2")
    st.caption("Parking inside the building is entered by floor in the Selected Floor tab.")

    st.divider()
    st.subheader("Loading Provided")
    loading_provided = st.number_input("Loading zones", min_value=0, step=1, key="loading_provided_v2")
    st.caption("Requirement: 1 loading zone per 100 units, rounded up.")

    st.divider()
    st.subheader("Saved Design Scenarios")
    scenario_name = st.text_input("Scenario name", placeholder="e.g. Scheme A")
    if st.button("Save Current Scenario", width="stretch"):
        clean_name = scenario_name.strip()
        if not clean_name:
            st.warning("Enter a scenario name first.")
        else:
            st.session_state.scenarios[clean_name] = snapshot_current()
            try:
                persist_scenarios()
                st.success(f"Saved: {clean_name}")
            except OSError as e:
                st.warning(f"Scenario saved for this session, but could not be written to disk: {e}")

    if st.session_state.scenarios:
        scenario_pick = st.selectbox("Saved scenarios", list(st.session_state.scenarios.keys()))
        c_load, c_delete = st.columns(2)
        if c_load.button("Load", width="stretch"):
            queue_snapshot_load(scenario_pick)
            st.rerun()
        if c_delete.button("Delete", width="stretch"):
            del st.session_state.scenarios[scenario_pick]
            try:
                persist_scenarios()
            except OSError:
                pass
            st.rerun()
    else:
        st.caption("No saved scenarios yet. Saved scenarios are stored on disk and survive browser refreshes.")


# ---------- Design Summary (top) ----------
unit_sizes = unit_size_dict(st.session_state.unit_sizes_df)
calc_top = recalc(st.session_state.floors, unit_sizes)

sum_gsf_top = calc_top["Total GSF"].sum()
sum_gfa_top = calc_top["GFA"].sum()
sum_lsf_top = calc_top["LSF"].sum()
total_units_top = int(calc_top["Units"].sum())
total_unit_sf_top = float(calc_top["Unit SF"].sum())
efficiency_top = (sum_lsf_top / sum_gsf_top) if sum_gsf_top else 0
far_top = (sum_gfa_top / lot_area) if lot_area else 0
allowable_gfa_top = lot_area * allowable_far
avg_unit_top = (total_unit_sf_top / total_units_top) if total_units_top else 0

studio_1br_top = int(calc_top[["S-A1", "1_Bd-B1", "1_Bd-B2", "1_Bd-B3", "1_Bd-B4"]].sum().sum())
two_3br_top = int(calc_top[["2_Bd-C1", "2_Bd-C2", "2_Bd-C3", "2_Bd-C4", "2_Bd-C5", "2_Bd-C6", "3_Bd-D1"]].sum().sum())
resident_parking_top = studio_1br_top * 1.0 + two_3br_top * 1.5
visitor_parking_top = total_units_top * 0.10
retail_parking_top = retail_sf * 3 / 1000
parking_required_top = resident_parking_top + visitor_parking_top + retail_parking_top
parking_in_building_top = int(calc_top["Parking Spaces"].sum())
parking_provided_top = p_offstreet + parking_in_building_top
loading_required_top = math.ceil(total_units_top / 100) if total_units_top else 0
loading_shortfall_top = max(loading_required_top - loading_provided, 0)

st.markdown("### Design Summary")
# Metric labels wrap instead of truncating with ellipses.
s1, s2, s3, s4, s5, s6 = st.columns(6)
s1.metric("Total GSF", f"{sum_gsf_top:,.0f} SF")
s2.metric("Total GFA", f"{sum_gfa_top:,.0f} SF")
s3.metric("Total LSF", f"{sum_lsf_top:,.0f} SF")
s4.metric("Efficiency", f"{efficiency_top:.1%}")
s5.metric("Total Units", f"{total_units_top:,}")
s6.metric("Avg Unit Size", f"{avg_unit_top:,.0f} SF")

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("FAR", f"{far_top:.2f}", f"Limit {allowable_far:.2f}", delta_color="off")
c2.metric("Allowable GFA", f"{allowable_gfa_top:,.0f} SF")
c3.metric("Parking Required", f"{parking_required_top:,.1f}")
c4.metric("Parking Provided", f"{parking_provided_top:,}", f"{parking_provided_top - parking_required_top:+.1f}")
c5.metric("Loading Required", f"{loading_required_top:,}")
c6.metric("Loading Provided", f"{loading_provided:,}", f"{loading_provided - loading_required_top:+d}")

status1, status2, status3 = st.columns(3)
with status1:
    if far_top <= allowable_far:
        st.success(f"FAR COMPLIES — {far_top:.2f} ≤ {allowable_far:.2f}")
    else:
        st.error(f"FAR DOES NOT COMPLY — {far_top:.2f} > {allowable_far:.2f}")
with status2:
    if parking_provided_top >= parking_required_top:
        st.success(f"PARKING COMPLIES — {parking_provided_top} provided vs. {parking_required_top:.1f} required")
    else:
        st.error(f"PARKING SHORTFALL — {parking_required_top - parking_provided_top:.1f} additional spaces required")
with status3:
    if loading_provided >= loading_required_top:
        st.success(f"LOADING COMPLIES — {loading_provided} provided vs. {loading_required_top} required")
    else:
        st.error(f"LOADING SHORTFALL — {loading_shortfall_top} additional loading zone{'s' if loading_shortfall_top != 1 else ''} required")

# ---------- Building controls ----------
st.markdown("### Building Controls")
control_left, control_right = st.columns([2, 3])
with control_left:
    floor_names = st.session_state.floors["Floor"].tolist()

    # Streamlit does not allow changing a widget's keyed session-state value
    # after that widget has been instantiated during the same run. Floor
    # actions therefore queue the next selection, rerun, and apply it here
    # before the selectbox is created.
    pending_floor = st.session_state.get("pending_selected_floor")
    if pending_floor in floor_names:
        st.session_state.selected_floor = pending_floor
        st.session_state.pending_selected_floor = None
    elif st.session_state.selected_floor not in floor_names:
        st.session_state.selected_floor = floor_names[0]

    selected_floor = st.selectbox("Selected floor", floor_names, key="selected_floor")
with control_right:
    st.caption("Add creates a blank level above the current highest floor. Duplicate inserts a copy directly above the selected floor, shifts every higher numbered floor up one level, and copies its unit mix and parking spaces. Delete automatically closes numbering gaps above Ground.")
    b1, b2, b3 = st.columns(3)
    if b1.button("＋ Add Floor", width="stretch"):
        new_row = make_blank_floor(st.session_state.floors)
        st.session_state.floors = sort_building(pd.concat([st.session_state.floors, pd.DataFrame([new_row])], ignore_index=True))
        st.session_state.pending_selected_floor = new_row["Floor"]
        bump_editor_keys()
        st.rerun()
    if b2.button("⧉ Duplicate Floor", width="stretch"):
        st.session_state.floors, new_name = duplicate_floor_in_sequence(
            st.session_state.floors, selected_floor
        )
        st.session_state.pending_selected_floor = new_name
        bump_editor_keys()
        st.rerun()
    if b3.button("🗑 Delete Floor", width="stretch"):
        if len(st.session_state.floors) <= 1:
            st.warning("The building must keep at least one floor.")
        else:
            st.session_state.floors = st.session_state.floors.loc[st.session_state.floors["Floor"] != selected_floor].reset_index(drop=True)
            # Close any numbering gap created by the deletion. This renames
            # every remaining numeric level consistently while leaving Ground
            # unchanged (e.g. deleting Parking Level 4 makes former Level 5
            # become Level 4, Level 6 become Level 5, and so on).
            st.session_state.floors = renumber_floors(st.session_state.floors)
            st.session_state.pending_selected_floor = st.session_state.floors.iloc[0]["Floor"]
            bump_editor_keys()
            st.rerun()

# ---------- Architectural tabs ----------
tab_stack, tab_areas, tab_mix, tab_sizes = st.tabs([
    "🏙 Building Stack", "📐 Floor Areas", "🏠 Selected Floor Unit Mix", "📏 Unit Types & Sizes"
])

with tab_stack:
    st.caption("The stack is proportional to each floor's Total GSF. Highest floor is shown first; Ground stays at the bottom.")
    render_building_stack(recalc(st.session_state.floors, unit_size_dict(st.session_state.unit_sizes_df)))

with tab_areas:
    st.caption("Edit floor names, uses, and area metrics here. Unit counts are handled separately in the Unit Mix tab.")
    area_cols = ["Floor", "Type", "Garage / Pool Deck", "Enclosed", "Balconies", "Total GSF", "LSF", "GFA", "Parking Spaces"]
    area_edited = st.data_editor(
        st.session_state.floors[area_cols],
        width="stretch",
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Floor": st.column_config.TextColumn("Floor", width="large", required=True),
            "Type": st.column_config.SelectboxColumn(
                "Type",
                options=["Residential", "Residential / Amenity", "Parking", "Ground / Podium", "Amenity", "Other"],
                required=True,
            ),
            "Garage / Pool Deck": st.column_config.NumberColumn("Garage / Pool Deck", min_value=0, format="%.0f"),
            "Enclosed": st.column_config.NumberColumn("Enclosed", min_value=0, format="%.0f"),
            "Balconies": st.column_config.NumberColumn("Balconies", min_value=0, format="%.0f"),
            "Total GSF": st.column_config.NumberColumn(
                "Total GSF", min_value=0, format="%.0f",
                help="Kept editable because the source workbook uses specific GSF values on some podium and amenity floors.",
            ),
            "LSF": st.column_config.NumberColumn("LSF", min_value=0, format="%.0f"),
            "GFA": st.column_config.NumberColumn("GFA", min_value=0, format="%.0f"),
            "Parking Spaces": st.column_config.NumberColumn("Parking Spaces", min_value=0, step=1, format="%d"),
        },
        key=f"floor_area_editor_{st.session_state.floor_editor_rev}",
    )
    # Preserve unit mixes while replacing the area columns.
    old_by_floor = st.session_state.floors.set_index("Floor")
    rebuilt = area_edited.copy()
    for u in UNIT_COLS:
        rebuilt[u] = rebuilt["Floor"].map(old_by_floor[u]).fillna(0).astype(int)
    # If a floor was renamed, preserve units by row position as a fallback.
    if len(rebuilt) == len(st.session_state.floors):
        for u in UNIT_COLS:
            missing_mask = rebuilt[u].isna() if rebuilt[u].dtype.kind == "f" else pd.Series(False, index=rebuilt.index)
            if missing_mask.any():
                rebuilt.loc[missing_mask, u] = st.session_state.floors.loc[missing_mask, u].values
    st.session_state.floors = sort_building(rebuilt)

with tab_mix:
    floor_names_now = st.session_state.floors["Floor"].tolist()
    if st.session_state.selected_floor not in floor_names_now:
        st.session_state.pending_selected_floor = floor_names_now[0]
        st.rerun()
    mix_floor = st.session_state.selected_floor
    floor_idx = st.session_state.floors.index[st.session_state.floors["Floor"] == mix_floor]
    if len(floor_idx):
        idx = floor_idx[0]
        mix_editor_df = st.session_state.unit_sizes_df.copy()
        mix_editor_df["Count"] = [int(st.session_state.floors.at[idx, u]) for u in mix_editor_df["Unit Type"]]
        mix_editor_df["Total Unit SF"] = mix_editor_df["Unit Size (SF)"] * mix_editor_df["Count"]
        st.markdown(f"#### {mix_floor}")
        st.caption("Edit unit counts below. Parking spaces for this floor can be entered here too — especially useful on parking levels.")
        parking_on_floor = st.number_input(
            "Parking spaces on this floor",
            min_value=0,
            step=1,
            value=int(st.session_state.floors.at[idx, "Parking Spaces"]),
            key=f"parking_floor_{st.session_state.mix_editor_rev}_{re.sub(r'[^A-Za-z0-9]+', '_', mix_floor)}",
        )
        st.session_state.floors.at[idx, "Parking Spaces"] = int(parking_on_floor)
        edited_mix = st.data_editor(
            mix_editor_df[["Category", "Unit Type", "Unit Size (SF)", "Count"]],
            width="stretch",
            hide_index=True,
            num_rows="fixed",
            disabled=["Category", "Unit Type", "Unit Size (SF)"],
            column_config={
                "Category": st.column_config.TextColumn("Category"),
                "Unit Type": st.column_config.TextColumn("Unit Type"),
                "Unit Size (SF)": st.column_config.NumberColumn("Size (SF)", format="%.0f"),
                "Count": st.column_config.NumberColumn("Count", min_value=0, step=1, format="%d"),
            },
            key=f"mix_editor_{st.session_state.mix_editor_rev}_{re.sub(r'[^A-Za-z0-9]+', '_', mix_floor)}",
        )
        for _, r in edited_mix.iterrows():
            st.session_state.floors.at[idx, r["Unit Type"]] = int(r["Count"])
        selected_calc = recalc(st.session_state.floors.loc[[idx]], unit_size_dict(st.session_state.unit_sizes_df)).iloc[0]
        u1, u2, u3, u4 = st.columns(4)
        u1.metric("Units on Floor", f"{int(selected_calc['Units'])}")
        u2.metric("Parking on Floor", f"{int(selected_calc['Parking Spaces'])}")
        u3.metric("Unit SF on Floor", f"{selected_calc['Unit SF']:,.0f} SF")
        u4.metric("LSF", f"{selected_calc['LSF']:,.0f} SF")

with tab_sizes:
    st.caption("Maintain the project-wide unit type library here. Changes update all unit-SF totals instantly.")
    edited_sizes = st.data_editor(
        st.session_state.unit_sizes_df,
        width="stretch",
        hide_index=True,
        num_rows="fixed",
        disabled=["Unit Type", "Category"],
        column_config={
            "Unit Type": st.column_config.TextColumn("Unit Type"),
            "Category": st.column_config.TextColumn("Category"),
            "Unit Size (SF)": st.column_config.NumberColumn("Unit Size (SF)", min_value=0, step=1, format="%d"),
        },
        key=f"unit_sizes_editor_{st.session_state.size_editor_rev}",
    )
    st.session_state.unit_sizes_df = edited_sizes.copy()


# ---------- Recalculate after all editors ----------
unit_sizes = unit_size_dict(st.session_state.unit_sizes_df)
calc = recalc(st.session_state.floors, unit_sizes)

sum_garage = calc["Garage / Pool Deck"].sum()
sum_enclosed = calc["Enclosed"].sum()
sum_balcony = calc["Balconies"].sum()
sum_gsf = calc["Total GSF"].sum()
sum_lsf = calc["LSF"].sum()
sum_gfa = calc["GFA"].sum()
total_units = int(calc["Units"].sum())
total_unit_sf = float(calc["Unit SF"].sum())
efficiency = (sum_lsf / sum_gsf) if sum_gsf else 0
far = (sum_gfa / lot_area) if lot_area else 0
allowable_gfa = lot_area * allowable_far
avg_unit = (total_unit_sf / total_units) if total_units else 0

studio_1br = int(calc[["S-A1", "1_Bd-B1", "1_Bd-B2", "1_Bd-B3", "1_Bd-B4"]].sum().sum())
two_3br = int(calc[["2_Bd-C1", "2_Bd-C2", "2_Bd-C3", "2_Bd-C4", "2_Bd-C5", "2_Bd-C6", "3_Bd-D1"]].sum().sum())
resident_parking = studio_1br * 1.0 + two_3br * 1.5
visitor_parking = total_units * 0.10
retail_parking = retail_sf * 3 / 1000
parking_required = resident_parking + visitor_parking + retail_parking
parking_in_building = int(calc["Parking Spaces"].sum())
parking_provided = p_offstreet + parking_in_building
loading_required = math.ceil(total_units / 100) if total_units else 0
loading_shortfall = max(loading_required - loading_provided, 0)

with st.expander("Calculated floor schedule"):
    display_cols = ["Floor", "Type", "Garage / Pool Deck", "Enclosed", "Balconies", "Total GSF", "LSF", "GFA", "Parking Spaces", "Units", "Unit SF"]
    st.dataframe(
        calc[display_cols].style.format({
            "Garage / Pool Deck": "{:,.0f}", "Enclosed": "{:,.0f}", "Balconies": "{:,.0f}",
            "Total GSF": "{:,.0f}", "LSF": "{:,.0f}", "GFA": "{:,.0f}", "Parking Spaces": "{:,.0f}", "Units": "{:,.0f}", "Unit SF": "{:,.0f}",
        }),
        width="stretch",
        hide_index=True,
    )

with st.expander("Parking calculation"):
    parking_df = pd.DataFrame([
        ["Studio + 1 BR", studio_1br, "× 1.0", studio_1br * 1.0],
        ["2 BR + 3 BR", two_3br, "× 1.5", two_3br * 1.5],
        ["Visitors", total_units, "× 10%", visitor_parking],
        ["Retail", retail_sf, "3 / 1,000 SF", retail_parking],
    ], columns=["Component", "Basis", "Rate", "Spaces"])
    st.dataframe(parking_df, hide_index=True, width="stretch")
    st.write(f"Building parking provided: **{parking_in_building:,}**")
    st.write(f"Off-street / site parking provided: **{p_offstreet:,}**")
    st.write(f"Total parking provided: **{parking_provided:,}**")


with st.expander("Loading calculation"):
    st.write(f"Total units: **{total_units:,}**")
    st.write(f"Requirement: **1 loading zone per 100 units, rounded up**")
    st.write(f"Loading zones required: **{loading_required}**")
    st.write(f"Loading zones provided: **{loading_provided}**")

with st.expander("Project unit mix totals"):
    mix = pd.DataFrame({
        "Unit Type": UNIT_COLS,
        "Unit Size (SF)": [unit_sizes[u] for u in UNIT_COLS],
        "Count": [int(calc[u].sum()) for u in UNIT_COLS],
    })
    mix["Total SF"] = mix["Unit Size (SF)"] * mix["Count"]
    st.dataframe(mix, hide_index=True, width="stretch")

# ---------- Excel export ----------
output = io.BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    calc.to_excel(writer, sheet_name="Floor Schedule", index=False)
    summary_export = pd.DataFrame({
        "Metric": [
            "Building Name", "Lot Area", "Allowable FAR", "Allowable GFA", "Total GSF", "Total GFA", "Total LSF",
            "Efficiency", "Total Units", "Average Unit Size", "Parking Required", "Parking Provided",
            "Loading Required", "Loading Provided"
        ],
        "Value": [
            building_name, lot_area, allowable_far, allowable_gfa, sum_gsf, sum_gfa, sum_lsf,
            efficiency, total_units, avg_unit, parking_required, parking_provided,
            loading_required, loading_provided
        ],
    })
    summary_export.to_excel(writer, sheet_name="Summary", index=False)
    mix.to_excel(writer, sheet_name="Unit Mix", index=False)
    st.session_state.unit_sizes_df.to_excel(writer, sheet_name="Unit Types", index=False)

st.download_button(
    "⬇ Download Current Design (.xlsx)",
    data=output.getvalue(),
    file_name="Floor_Area_Design.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption("V2.3.0 applies the Stormy Morning interface palette across the app while preserving all V2.2.4 functionality.")
