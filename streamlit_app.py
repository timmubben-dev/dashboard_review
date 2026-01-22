import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime

# Web-Oberfläche Design
st.set_page_config(page_title="Herzklappen Dashboard Generator", layout="wide")
st.title("🏥 Herzklappen Dashboard Generator")
st.markdown("""
Dieses Tool erstellt ein automatisiertes Master-Dashboard aus Ihrer Klappen-Strukturliste. 
Laden Sie einfach Ihre Excel-Datei hoch (der Dateiname ist egal).
""")

# Datei-Uploader (Akzeptiert jede .xlsx Datei)
uploaded_file = st.file_uploader("Excel-Datei hier hineinziehen oder klicken", type=["xlsx"])

def map_to_kpi(e):
    e = str(e).lower()
    if 'tavi' in e: return 'TAVI'
    if 'edge-to-edge mk' in e or 'tmvi' in e: return 'MTEER'
    if 'edge-to-edge tk' in e or 'htp tk' in e: return 'TTEER'
    if 'ttvi' in e: return 'TTVI'
    if 'tricvalve' in e or 'ttvr' in e: return 'TTVR'
    return 'Sonstige'

if uploaded_file:
    try:
        # Daten einlesen (Name der Datei ist hier egal)
        df = pd.read_excel(uploaded_file, sheet_name='Daten', skiprows=5, engine='openpyxl')
        df = df[df['Nr.'].notnull()].copy()
        
        # Berechnungen
        df['Prozedur_Date'] = pd.to_datetime(df['Prozedur'], errors='coerce')
        df['Year'] = df['Prozedur_Date'].dt.year
        df['Month'] = df['Prozedur_Date'].dt.month
        df['KPI_Kat'] = df['Eingriff'].apply(map_to_kpi)
        df['VWD_num'] = pd.to_numeric(df['VWD'], errors='coerce')
        
        df_2026 = df[df['Year'] == 2026].copy()
        months_passed = df_2026['Month'].max() or 1
        heute_str = datetime.now().strftime('%d-%m-%Y')

        # Excel-Erstellung im Speicher
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            ws = workbook.add_worksheet('Master Dashboard')
            
            # Formate definieren
            title_f = workbook.add_format({'bold': True, 'size': 14, 'font_color': '#1F4E78', 'bottom': 2})
            date_f = workbook.add_format({'bold': True, 'align': 'right', 'font_color': '#595959'})
            header_f = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
            cell_f = workbook.add_format({'border': 1, 'align': 'center'})
            pct_f = workbook.add_format({'num_format': '0.0%', 'border': 1, 'align': 'center'})
            num_f = workbook.add_format({'num_format': '0.0', 'border': 1, 'align': 'center'})
            red_f = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center'})
            green_f = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center'})
            yellow_pct_f = workbook.add_format({'bg_color': '#FFEB9C', 'border': 1, 'num_format': '0.0%', 'align': 'center'})
            green_pct_f = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'num_format': '0.0%', 'align': 'center'})
            red_pct_f = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'num_format': '0.0%', 'align': 'center'})

            # 1. LEISTUNGSZAHLEN
            ws.write('A1', '1. LEISTUNGSZAHLEN & PROGNOSE 2026', title_f)
            ws.write('Q1', f'Stand: {heute_str}', date_f)
            targets = {'TAVI': 46, 'MTEER': 10, 'TTEER': 7, 'TTVI': 3, 'TTVR': 1}
            headers = ['Kategorie'] + ['Jan','Feb','Mrz','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Dez'] + ['YTD','Prognose','Soll','Status']
            for c, h in enumerate(headers): ws.write(2, c, h, header_f)
            for r, (cat, t_mo) in enumerate(targets.items()):
                cat_df = df_2026[df_2026['KPI_Kat'] == cat]
                counts = cat_df.groupby('Month').size()
                ws.write(r+3, 0, cat, workbook.add_format({'bold': True, 'border': 1}))
                for m in range(1, 13):
                    val = counts.get(m, 0)
                    fmt = red_f if (val < t_mo and m <= months_passed) else cell_f
                    ws.write(r+3, m, val, fmt)
                ist_ytd = len(cat_df); fc = round((ist_ytd / months_passed) * 12)
                ws.write(r+3, 13, ist_ytd, cell_f); ws.write(r+3, 14, fc, cell_f)
                ws.write(r+3, 15, t_mo * 12, cell_f)
                ws.write_formula(r+3, 16, f'=IFERROR(O{r+4}/P{r+4}, 0)', pct_f)

            # 2. VERWEILDAUER (AB 2024)
            ws.write('A10', '2. VERWEILDAUER (ZIEL: 5 TAGE MEDIAN)', title_f)
            for c, h in enumerate(['Jahr', 'VWD Alle (Med)', 'VWD <21d (Mittel)', 'VWD <21d (Med)']): ws.write(11, c, h, header_f)
            for r, y in enumerate([2024, 2025, 2026]):
                y_df = df[df['Year'] == y]
                v_short = y_df[(y_df['VWD_num'] > 0) & (y_df['VWD_num'] < 21)]['VWD_num']
                ws.write(12+r, 0, y, cell_f)
                ws.write(12+r, 1, y_df[(y_df['VWD_num'] > 0)]['VWD_num'].median() if not y_df.empty else 0, cell_f)
                ws.write(12+r, 2, v_short.mean() if not v_short.empty else 0, num_f)
                ms = v_short.median() if not v_short.empty else 0
                ws.write(12+r, 3, ms, green_f if 0 < ms <= 5 else red_f)

            # 4. TAVI-TEAMS
            ws.write('A22', '4. TAVI-TEAMS 2026', title_f)
            tavi_2026 = df_2026[df_2026['KPI_Kat'] == 'TAVI']
            team_stats = tavi_2026['Team'].value_counts().reset_index()
            ws.write(23, 0, 'Team', header_f); ws.write(23, 1, 'Fälle', header_f); ws.write(23, 2, 'Anteil (%)', header_f)
            for r, row in enumerate(team_stats.values):
                ws.write(24+r, 0, row[0], cell_f); ws.write(24+r, 1, row[1], cell_f)
                ws.write(24+r, 2, row[1]/len(tavi_2026) if len(tavi_2026) > 0 else 0, pct_f)

            # 5. STRATEGIE & QUALITÄT
            ws.write('A32', '5. STRATEGIE & QUALITÄT 2026', title_f)
            ev_r = tavi_2026['Device'].str.contains('Evolut', na=False, case=False).mean() if len(tavi_2026) > 0 else 0
            ws.write(33, 0, 'Evolut-Anteil (Ziel 80%)', cell_f)
            ws.write(33, 1, ev_r, green_pct_f if ev_r >= 0.8 else red_pct_f if ev_r < 0.7 else yellow_pct_f)

            # 6. HISTORIE
            ws.write('A39', '6. HISTORISCHE ENTWICKLUNG', title_f)
            h_cats = ['TAVI', 'MTEER', 'TTEER']
            for c, h in enumerate(['Jahr'] + h_cats): ws.write(40, c, h, header_f)
            for r, y in enumerate([2022, 2023, 2024, 2025, 2026]):
                ws.write(41+r, 0, y, cell_f)
                for i, cat in enumerate(h_cats):
                    ws.write(41+r, i+1, len(df[(df['Year'] == y) & (df['KPI_Kat'] == cat)]), cell_f)

            ws.set_column('A:A', 36); ws.set_column('B:M', 6.8); ws.set_column('N:Q', 13)

        st.success("✅ Dashboard erfolgreich generiert!")
        st.download_button(
            label="📊 Dashboard herunterladen",
            data=output.getvalue(),
            file_name=f"Dashboard_{heute_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Fehler: {e}. Stellen Sie sicher, dass das Tabellenblatt 'Daten' heißt.")
