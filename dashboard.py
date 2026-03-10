import os
from io import BytesIO

import pandas as pd
import streamlit as st
import altair as alt

st.set_page_config(layout="wide")

# Altair: lichte grafieken (achtergrond + assen/legenda donkere tekst)
def _light_chart(chart):
    return (
        chart.properties(background="#f5f5f5")
        .configure_axis(labelColor="#374151", titleColor="#374151", gridColor="#e5e7eb", domainColor="#9ca3af")
        .configure_legend(labelColor="#374151", titleColor="#374151")
    )

# Minimalist wit/grijs thema
st.markdown("""
<style>
    .stApp { background-color: #fafafa; }
    h1, h2, h3 { color: #374151 !important; font-weight: 500 !important; }
    [data-testid="stMetricValue"] { color: #4b5563; font-weight: 500; }
    [data-testid="stMetricLabel"] { color: #6b7280; }
    
    /* KPI-boxen */
    [data-testid="stMetric"] {
        background-color: #f5f5f5 !important;
        border: 1px dashed #d1d5db;
        border-radius: 8px;
        padding: 1rem;
    }
    
    /* Text input – getypte tekst zwart, geen rode hover/focus */
    [data-testid="stTextInput"] input, [data-testid="stTextInput"] input::placeholder {
        background-color: #f5f5f5 !important;
        border: 1px dashed #d1d5db !important;
        border-radius: 6px;
        color: #111827 !important;
    }
    [data-testid="stTextInput"] input:focus, [data-testid="stTextInput"] input:hover {
        border-color: #9ca3af !important;
        box-shadow: none !important;
    }
    input[type="password"], input[type="text"] {
        color: #111827 !important;
    }
    
    /* Grafieken: lichte container */
    [data-testid="stArrowVegaLiteChart"], [data-testid="stArrowAltairChart"],
    [data-testid="stVegaLiteChart"] {
        background-color: #f5f5f5 !important;
        border: 1px dashed #d1d5db;
        border-radius: 8px;
        padding: 1rem;
    }
    
    /* Dataframe wrapper: grijze rand */
    .stDataFrame, [data-testid="stDataFrame"] {
        border: 1px dashed #d1d5db;
        border-radius: 8px;
    }
    
    /* Buttons: geen rode hover, alleen grijze rand */
    [data-testid="stButton"] button:hover, [data-testid="stButton"] button:focus,
    button[kind="primary"]:hover, button[kind="primary"]:focus {
        border-color: #9ca3af !important;
        box-shadow: none !important;
    }
    
    hr { border-color: #e5e7eb !important; }
    
    /* Sidebar: smaller */
    [data-testid="stSidebar"] { background-color: #f5f5f5; width: 12rem !important; min-width: 12rem !important; }
    
    /* Uitloggen-knop: grijs met zwarte letters */
    [data-testid="stSidebar"] [data-testid="stButton"] button {
        background-color: #f5f5f5 !important;
        color: #111827 !important;
        border: 1px dashed #d1d5db !important;
    }
    
    /* Inlogformulier: gecentreerd en groter */
    .st-key-login-box input, [class*="login-box"] input {
        font-size: 1.15rem !important; padding: 0.75rem 1rem !important;
    }
    .st-key-login-box button, [class*="login-box"] button {
        font-size: 1.15rem !important; padding: 0.65rem 2rem !important; min-height: 2.75rem;
    }
    .st-key-login-box label, [class*="login-box"] label {
        font-size: 1.1rem !important;
    }
</style>
""", unsafe_allow_html=True)

# Inloggen
ALLOWED_NAMES = {"rob", "stef", "frank", "edwin", "franklin", "berthil", "twan"}
PASSWORD = os.getenv("DASHBOARD_PASSWORD", "")

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.markdown('<h1 style="text-align: center; color: #374151; font-weight: 500;">Klanten omzetanalyse – Inloggen</h1>', unsafe_allow_html=True)
    if not PASSWORD:
        st.error("DASHBOARD_PASSWORD is niet geconfigureerd. Stel deze omgevingsvariabele in (bijv. op Railway).")
        st.stop()
    _, col_form, _ = st.columns([1, 3, 1])
    with col_form:
        with st.container(key="login-box"):
            naam = st.text_input("Naam", key="login_naam")
            wachtwoord = st.text_input("Wachtwoord", type="password", key="login_pw")
            if st.button("Inloggen", key="login_btn"):
                if naam.strip().lower() in ALLOWED_NAMES and wachtwoord == PASSWORD:
                    st.session_state.logged_in = True
                    st.rerun()
                else:
                    st.error("Ongeldige naam of wachtwoord.")
    st.stop()

with st.sidebar:
    if st.button("Uitloggen"):
        st.session_state.logged_in = False
        st.rerun()

# Data laden: lokaal bestand of upload
LOCAL_FILES = [
    os.getenv("LOCAL_DATA_FILE"),
    "klanten_omzet_analyse.xlsx",
    "klanten_met_omzetdaling.xlsx",
    "Omzet 2025 klanten.xlsm",
]

def _load_raw_excel(file) -> pd.DataFrame:
    """Laadt ruwe Excel (Omzet 2025-formaat) en retourneert merged dataframe."""
    df = pd.read_excel(file, sheet_name=0, header=None)
    header_row = 13
    data = df.iloc[header_row:].copy()
    data.columns = data.iloc[0]
    data = data[1:]
    left = data.iloc[:, 0:4]
    left.columns = ["Relatiecode", "Relatienaam", "Omzet_A", "Marge_A"]
    right = data.iloc[:, 5:10]
    right.columns = ["Relatiecode", "Relatienaam2", "Details", "Omzet_B", "Marge_B"]
    right = right[["Relatiecode", "Relatienaam2", "Omzet_B", "Marge_B"]]
    merged = pd.merge(left, right, on="Relatiecode", how="outer")
    merged = merged[
        ~merged["Relatienaam"].astype(str).str.contains(
            "totaal|total|grand", case=False, na=False
        )
    ]
    merged = merged.reset_index(drop=True)
    merged["Naam"] = merged["Relatienaam"].fillna(merged["Relatienaam2"])
    merged = merged[merged["Naam"].notna()]
    return merged

def _normalize_processed(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliseert verwerkt Excel/CSV naar dashboard-formaat."""
    df = df.copy()
    # Zorg dat Naam altijd bestaat
    if "Naam" not in df.columns:
        for col in ["Relatienaam", "relatienaam", "Naam", "naam", "Relatienaam2"]:
            if col in df.columns:
                df["Naam"] = df[col]
                break
        if "Naam" not in df.columns:
            df["Naam"] = df.iloc[:, 1] if len(df.columns) > 1 else ""
    if "% verandering" in df.columns and "Procent_verandering" not in df.columns:
        df["Procent_verandering"] = pd.to_numeric(df["% verandering"], errors="coerce").fillna(0)
    if "Omzet_verschil" not in df.columns:
        df["Omzet_verschil"] = df["Omzet_B"] - df["Omzet_A"]
    if "Procent_verandering" not in df.columns:
        df["Procent_verandering"] = 0
        mask = df["Omzet_A"] != 0
        df.loc[mask, "Procent_verandering"] = (
            df.loc[mask, "Omzet_verschil"] / df.loc[mask, "Omzet_A"] * 100
        )
    return df

def _is_processed_format(df: pd.DataFrame) -> bool:
    """Controleert of dit al verwerkte data is (heeft Omzet_A, Omzet_B)."""
    cols = [c for c in df.columns if isinstance(c, str)]
    return "Omzet_A" in cols and "Omzet_B" in cols and "Relatiecode" in cols

# Bepaal bron: lokaal bestand
file_obj = None
for path in LOCAL_FILES:
    if path and os.path.isfile(path):
        file_obj = path
        break

if not file_obj:
    st.warning("Geen bestand gevonden. Zet een Excel of CSV in de projectmap (bijv. klanten_omzet_analyse.xlsx).")
    st.stop()

try:
    is_csv = getattr(file_obj, "name", str(file_obj)).lower().endswith(".csv")
    if is_csv:
        df = pd.read_csv(file_obj)
        merged = _normalize_processed(df) if _is_processed_format(df) else None
        if merged is None:
            st.error("CSV moet kolommen hebben: Relatiecode, Omzet_A, Omzet_B (of Relatienaam)")
            st.stop()
    else:
        df = pd.read_excel(file_obj, sheet_name=0)
        if _is_processed_format(df):
            merged = _normalize_processed(df)
        else:
            merged = _load_raw_excel(file_obj)
except FileNotFoundError:
    st.warning("Lokaal bestand niet gevonden. Zet LOCAL_DATA_FILE of plaats het bestand in de projectmap.")
    st.stop()
except Exception as e:
    st.error(f"Fout bij laden: {e}")
    st.stop()

# types corrigeren
merged["Relatiecode"] = merged["Relatiecode"].astype(str)
if "Naam" not in merged.columns:
    name_col = None
    for col in merged.columns:
        if col and str(col).strip().lower() in ("relatienaam", "naam", "relatienaam2"):
            name_col = col
            break
    merged["Naam"] = merged[name_col] if name_col else (merged.iloc[:, 1] if len(merged.columns) > 1 else pd.Series([""] * len(merged), index=merged.index))
merged["Naam"] = merged["Naam"].fillna("").astype(str)

# Totaalregels verwijderen
merged = merged[
    ~merged["Naam"].str.contains("totaal|total|grand", case=False, na=False)
]

merged["Omzet_A"] = pd.to_numeric(merged["Omzet_A"], errors="coerce").fillna(0)
merged["Omzet_B"] = pd.to_numeric(merged["Omzet_B"], errors="coerce").fillna(0)

# Specifieke totalerij uitsluiten (omzet rond 1600991.94)
merged = merged[
    ~((merged["Omzet_A"].round(2) == 1600991.94) | (merged["Omzet_B"].round(2) == 1600991.94))
]

# Berekeningen
merged["Omzet_verschil"] = merged["Omzet_B"] - merged["Omzet_A"]

merged["Procent_verandering"] = 0
mask = merged["Omzet_A"] != 0

merged.loc[mask,"Procent_verandering"] = (
    merged.loc[mask,"Omzet_verschil"] /
    merged.loc[mask,"Omzet_A"] * 100
)

# Opportunity score
merged["Opportunity_score"] = abs(merged["Omzet_verschil"])

# Klant classificatie
merged["Klant_type"] = "Stabiele klant"

merged.loc[
    (merged["Omzet_A"] > 5000) &
    (merged["Procent_verandering"] < -20),
    "Klant_type"
] = "Churn risico"

merged.loc[
    (merged["Omzet_A"] > 5000) &
    (merged["Omzet_B"] < 1000),
    "Klant_type"
] = "Heractivatie kans"

merged.loc[
    (merged["Procent_verandering"] > 20) &
    (merged["Omzet_B"] > merged["Omzet_A"]),
    "Klant_type"
] = "Upsell kans"

st.title("Klanten omzetanalyse")

# KPI blok
col1,col2,col3,col4 = st.columns(4)

col1.metric("Totale omzet vorig jaar", f"€{merged['Omzet_A'].sum():,.0f}")
col2.metric("Totale omzet huidig", f"€{merged['Omzet_B'].sum():,.0f}")
totaal_verschil = merged["Omzet_verschil"].sum()
col3.metric("Totaal verschil", f"€{totaal_verschil:+,.0f}")
col4.metric("Aantal klanten", len(merged))

st.divider()

# Segment overzicht
seg = merged["Klant_type"].value_counts()

col1,col2,col3 = st.columns(3)

col1.metric("Churn risico", seg.get("Churn risico",0))
col2.metric("Upsell kansen", seg.get("Upsell kans",0))
col3.metric("Heractivatie kansen", seg.get("Heractivatie kans",0))

st.divider()

# Zoekfunctie
zoek = st.text_input("Zoek klantnummer of klantnaam")

if zoek:

    result = merged[
        merged["Naam"].astype(str).str.contains(zoek, case=False, na=False) |
        merged["Relatiecode"].astype(str).str.contains(zoek)
    ]

    if len(result)>0:

        klant = result.iloc[0]

        st.subheader("Klantprofiel")
        st.markdown(f"**Bedrijfsnaam:** {klant['Naam']}  \n**Klantnummer:** {klant['Relatiecode']}")
        st.divider()

        col1,col2,col3,col4,col5 = st.columns(5)

        col1.metric("Omzet vorig jaar", f"€{klant['Omzet_A']:,.0f}")
        col2.metric("Omzet huidig", f"€{klant['Omzet_B']:,.0f}")
        col3.metric("Verschil", f"€{klant['Omzet_verschil']:,.0f}")
        col4.metric("Verandering", f"{klant['Procent_verandering']:.1f}%")
        col5.metric("Segment", klant["Klant_type"])

        klant_df = pd.DataFrame({
            "Periode":["Vorig jaar","Huidig"],
            "Omzet":[klant["Omzet_A"],klant["Omzet_B"]]
        })

        chart = alt.Chart(klant_df).mark_bar().encode(
            x="Periode",
            y="Omzet",
            color=alt.value("#6b7280")
        )
        st.altair_chart(_light_chart(chart), use_container_width=True)

        # Export klantprofiel
        klant_export = pd.DataFrame([{
            "Relatiecode": klant["Relatiecode"],
            "Naam": klant["Naam"],
            "Omzet vorig jaar": klant["Omzet_A"],
            "Omzet huidig": klant["Omzet_B"],
            "Verschil": klant["Omzet_verschil"],
            "Procent verandering": f"{klant['Procent_verandering']:.1f}%",
            "Segment": klant["Klant_type"],
        }])
        col_csv, col_excel, _ = st.columns([1, 1, 2])
        with col_csv:
            st.download_button(
                "Exporteer CSV",
                data=klant_export.to_csv(index=False),
                file_name=f"klantprofiel_{klant['Relatiecode']}.csv",
                mime="text/csv",
                key=f"export_csv_{klant['Relatiecode']}",
            )
        with col_excel:
            buffer = BytesIO()
            klant_export.to_excel(buffer, index=False, engine="openpyxl")
            st.download_button(
                "Exporteer Excel",
                data=buffer.getvalue(),
                file_name=f"klantprofiel_{klant['Relatiecode']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"export_excel_{klant['Relatiecode']}",
            )

st.divider()

# Grootste dalers
st.subheader("Grootste omzetdalers")

top_dalers = merged.nsmallest(40,"Omzet_verschil")
chart = alt.Chart(top_dalers).mark_bar().encode(
    x="Omzet_verschil",
    y=alt.Y("Naam", type="nominal", sort="x"),
    color=alt.value("#dc2626"),
    tooltip=["Naam","Relatiecode","Omzet_verschil"]
)
st.altair_chart(_light_chart(chart), use_container_width=True)

st.divider()
st.markdown("<div style='height: 3rem;'></div>", unsafe_allow_html=True)

# Top sales kansen
st.subheader("Top sales kansen")

kansen = merged.nlargest(20,"Opportunity_score")
st.dataframe(
    kansen[["Relatiecode","Naam","Omzet_A","Omzet_B","Omzet_verschil","Klant_type"]],
    use_container_width=True
)

st.divider()

# Groei grafiek
st.subheader("Grootste omzetgroei")

groei = merged.nlargest(40,"Omzet_verschil")

chart2 = alt.Chart(groei).mark_bar().encode(
    x="Omzet_verschil",
    y=alt.Y("Naam", type="nominal", sort="-x"),
    color=alt.value("#16a34a"),
    tooltip=["Naam","Relatiecode","Omzet_verschil"]
)
st.altair_chart(_light_chart(chart2), use_container_width=True)

st.divider()

# Opportunity matrix
st.subheader("Opportunity matrix")

matrix = alt.Chart(merged).mark_circle(size=120).encode(
    x=alt.X("Omzet_A",title="Omzet vorig jaar"),
    y=alt.Y("Omzet_verschil",title="Omzetverandering"),
    color=alt.Color("Klant_type", scale=alt.Scale(
        domain=["Churn risico", "Heractivatie kans", "Stabiele klant", "Upsell kans"],
        range=["#dc2626", "#ea580c", "#6b7280", "#16a34a"]
    )),
    tooltip=[
        "Naam",
        "Relatiecode",
        "Omzet_A",
        "Omzet_B",
        "Omzet_verschil",
        "Procent_verandering",
        "Klant_type"
    ]
).interactive()
st.altair_chart(_light_chart(matrix), use_container_width=True)

st.divider()

# Klant tabel
st.subheader("Alle klanten")

st.dataframe(
    merged[["Relatiecode","Naam","Omzet_A","Omzet_B","Omzet_verschil","Procent_verandering","Klant_type"]],
    use_container_width=True
)