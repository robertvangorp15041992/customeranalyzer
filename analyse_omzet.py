import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.rule import ColorScaleRule

file = "Omzet 2025 klanten.xlsm"

df = pd.read_excel(file, sheet_name=0, header=None)

header_row = 13
data = df.iloc[header_row:].copy()
data.columns = data.iloc[0]
data = data[1:]

left = data.iloc[:,0:4]
left.columns = ["Relatiecode","Relatienaam","Omzet_A","Marge_A"]

right = data.iloc[:,5:10]
right.columns = ["Relatiecode","Relatienaam","Details","Omzet_B","Marge_B"]
right = right[["Relatiecode","Relatienaam","Omzet_B","Marge_B"]]

merged = pd.merge(left, right, on="Relatiecode", how="outer")

merged["Omzet_A"] = pd.to_numeric(merged["Omzet_A"], errors="coerce")
merged["Omzet_B"] = pd.to_numeric(merged["Omzet_B"], errors="coerce")

merged["Omzet_verschil"] = merged["Omzet_B"] - merged["Omzet_A"]
merged["% verandering"] = (merged["Omzet_verschil"] / merged["Omzet_A"]) * 100

merged = merged.sort_values("Omzet_verschil")

output_file = "klanten_omzet_analyse.xlsx"
merged.to_excel(output_file, index=False)

# Excel openen om kleurregels toe te voegen
wb = load_workbook(output_file)
ws = wb.active

# Kleurverloop regel voor kolom verschil (E)
rule = ColorScaleRule(
    start_type='min',
    start_color='F8696B',   # rood
    mid_type='percentile',
    mid_value=50,
    mid_color='FFEB84',     # geel
    end_type='max',
    end_color='63BE7B'      # groen
)

ws.conditional_formatting.add("E2:E1000", rule)

wb.save(output_file)

print("Analyse compleet met kleurcodes:", output_file)