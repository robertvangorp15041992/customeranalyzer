import pandas as pd
import streamlit as st
import altair as alt

st.set_page_config(layout="wide")

# Excel laden: upload of lokaal bestand
uploaded = st.file_uploader("Upload omzet Excel (of gebruik lokaal bestand)", type=["xlsx", "xlsm", "xls"])
if uploaded:
    file = uploaded
else:
    file = "Omzet 2025 klanten.xlsm"

try:
    df = pd.read_excel(file, sheet_name=0, header=None)
except FileNotFoundError:
    st.warning("Geen bestand gevonden. Upload een omzet Excel-bestand (.xlsx of .xlsm) om te beginnen.")
    st.stop()
except Exception as e:
    st.error(f"Fout bij laden: {e}")
    st.stop()

header_row = 13
data = df.iloc[header_row:].copy()
data.columns = data.iloc[0]
data = data[1:]

# Periode A
left = data.iloc[:,0:4]
left.columns = ["Relatiecode","Relatienaam","Omzet_A","Marge_A"]

# Periode B
right = data.iloc[:,5:10]
right.columns = ["Relatiecode","Relatienaam2","Details","Omzet_B","Marge_B"]
right = right[["Relatiecode","Relatienaam2","Omzet_B","Marge_B"]]

# Merge
merged = pd.merge(left, right, on="Relatiecode", how="outer")

# Totaalregels verwijderen
merged = merged[
    ~merged["Relatienaam"].astype(str).str.contains(
        "totaal|total|grand", case=False, na=False
    )
]

merged = merged.reset_index(drop=True)

# Naam samenvoegen
merged["Naam"] = merged["Relatienaam"].fillna(merged["Relatienaam2"])

# lege klanten verwijderen
merged = merged[merged["Naam"].notna()]

# types corrigeren
merged["Relatiecode"] = merged["Relatiecode"].astype(str)

merged["Omzet_A"] = pd.to_numeric(merged["Omzet_A"], errors="coerce").fillna(0)
merged["Omzet_B"] = pd.to_numeric(merged["Omzet_B"], errors="coerce").fillna(0)

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
col3.metric("Totaal verschil", f"€{merged['Omzet_verschil'].sum():,.0f}")
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
zoek = st.text_input("Zoek relatiecode of bedrijfsnaam")

if zoek:

    result = merged[
        merged["Naam"].astype(str).str.contains(zoek, case=False, na=False) |
        merged["Relatiecode"].astype(str).str.contains(zoek)
    ]

    if len(result)>0:

        klant = result.iloc[0]

        st.subheader("Klantprofiel")

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
            color=alt.value("#4c78a8")
        )

        st.altair_chart(chart,use_container_width=True)

st.divider()

# Grootste dalers
col1,col2 = st.columns(2)

with col1:

    st.subheader("Grootste omzetdalers")

    top_dalers = merged.nsmallest(20,"Omzet_verschil")

    chart = alt.Chart(top_dalers).mark_bar().encode(
        x="Omzet_verschil",
        y=alt.Y("Naam",sort="x"),
        color=alt.value("#c0392b"),
        tooltip=["Naam","Relatiecode","Omzet_verschil"]
    )

    st.altair_chart(chart,use_container_width=True)

with col2:

    st.subheader("Top sales kansen")

    kansen = merged.nlargest(20,"Opportunity_score")

    st.dataframe(
        kansen[
            ["Relatiecode","Naam","Omzet_A","Omzet_B","Omzet_verschil","Klant_type"]
        ],
        use_container_width=True
    )

st.divider()

# Groei grafiek
st.subheader("Grootste omzetgroei")

groei = merged.nlargest(20,"Omzet_verschil")

chart2 = alt.Chart(groei).mark_bar().encode(
    x="Omzet_verschil",
    y=alt.Y("Naam",sort="-x"),
    color=alt.value("#1e8449"),
    tooltip=["Naam","Relatiecode","Omzet_verschil"]
)

st.altair_chart(chart2,use_container_width=True)

st.divider()

# Opportunity matrix
st.subheader("Opportunity matrix")

matrix = alt.Chart(merged).mark_circle(size=120).encode(
    x=alt.X("Omzet_A",title="Omzet vorig jaar"),
    y=alt.Y("Omzet_verschil",title="Omzetverandering"),
    color="Klant_type",
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

st.altair_chart(matrix,use_container_width=True)

st.divider()

# Klant tabel
st.subheader("Alle klanten")

st.dataframe(
    merged[
        ["Relatiecode","Naam","Omzet_A","Omzet_B","Omzet_verschil","Procent_verandering","Klant_type"]
    ],
    use_container_width=True
)