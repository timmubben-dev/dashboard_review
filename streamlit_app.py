import streamlit as st
import pandas as pd
import numpy as np
import io
import msoffcrypto
from datetime import datetime

# 1. SEITEN-KONFIGURATION
st.set_page_config(page_title="Herzklappen Master-Dashboard", layout="wide")
st.title("🏥 Herzklappen Master-Dashboard 2026")

# Sidebar für Login & Upload
st.sidebar.header("🔐 Datensicherheit")
password = st.sidebar.text_input("Bitte Excel-Passwort eingeben", type="password")
uploaded_file = st.sidebar.file_uploader("Verschlüsselte Excel-Datei wählen", type=["xlsx"])

def map_to_kpi(e):
    if pd.isna(e): return 'Sonstige'
    e = str(e).lower().strip()
    if 'tavi' in e: return 'TAVI'
    if any(x in e for x in ['edge-to-edge mk', 'tmvi', 'mteer']): return 'MTEER'
    if any(x in e for x in ['edge-to-edge tk', 'htp tk', 'tteer']): return 'TTEER'
    if 'ttvi' in e: return 'TTVI'
    if any(x in e for x in ['tricvalve', 'ttvr']): return 'TTVR'
    return 'Sonstige'

if uploaded_file and password:
    try:
        # --- DATEI ENTSCHLÜSSELN ---
        decrypted_file = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(uploaded_file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted_file)
        decrypted_file.seek(0)

        # --- DATEN LADEN ---
        df = pd.read_excel(decrypted_file, sheet_name='Daten', skiprows=5, engine='openpyxl')
        df = df[df['Nr.'].notnull()].copy()
        
        df['Prozedur_Date'] = pd.to_datetime(df['Prozedur'], errors='coerce')
        df['Year'] = df['Prozedur_Date'].dt.year
        df['Month'] = df['Prozedur_Date'].dt.month
        df['KPI_Kat'] = df['Eingriff'].apply(map_to_kpi)
        df['VWD_num'] = pd.to_numeric(df['VWD'], errors='coerce')
        
        df_2026 = df[df['Year'] == 2026].copy()
        months_passed = int(df_2026['Month'].max()) if not df_2026.empty else 1
        heute_str = datetime.now().strftime('%d-%m-%Y')

        # --- EXCEL GENERIERUNG ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            ws = workbook.add_worksheet('Master Dashboard')
            
            # ALLE FORMATE VORAB DEFINIEREN
            f_title = workbook.add_format({'bold': True, 'size': 12, 'font_color': '#1F4E78', 'bottom': 2})
            f_header = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
            f_cell = workbook.add_format({'border': 1, 'align': 'center'})
            f_bold_cell = workbook.add_format({'bold': True, 'border': 1, 'align': 'left'})
            f_pct = workbook.add_format({'num_format': '0.0%', 'border': 1, 'align': 'center'})
            f_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center'})
            f_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center'})
            f_num = workbook.add_format({'num_format': '0.0', 'border': 1, 'align': 'center'})

            # --- 1. LEISTUNGSZAHLEN ---
            ws.write(0, 0, '1. LEISTUNGSZAHLEN & PROGNOSE 2026', f_title)
            targets = {'TAVI': 46, 'MTEER': 10, 'TTEER': 7, 'TTVI': 3, 'TTVR': 1}
            headers = ['Kategorie', 'Jan','Feb','Mrz','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Dez', 'YTD','Prognose','Soll','Status']
            for c, h in enumerate(headers):
                ws.write(2, c, h, f_header)
            
            for r, (cat, t_mo) in enumerate(targets.items()):
                c_df = df_2026[df_2026['KPI_Kat'] == cat]
                counts = c_df.groupby('Month').size()
                ws.write(r+3, 0, cat, f_bold_cell)
                for m in range(1, 13):
                    val = int(counts.get(m, 0))
                    fmt = f_red if (val < t_mo and m <= months_passed) else f_cell
                    ws.write(r+3, m, val, fmt)
                
                ist_y = len(c_df)
                prog = round((ist_y / months_passed) * 12) if months_passed > 0 else 0
                ws.write(r+3, 13, ist_y, f_cell)
                ws.write(r+3, 14, prog, f_cell)
                ws.write(r+3, 15, t_mo * 12, f_cell)
                ws.write_formula(r+3, 16, f'=IFERROR(O{r+4}/P{r+4}, 0)', f_pct)

            # --- 2. VERWEILDAUER ---
            curr_r = 10
            ws.write(curr_r, 0, '2. VERWEILDAUER (ZIEL: 5 TAGE MEDIAN)', f_title)
            vwd_h = ['Jahr', 'VWD Alle (Med)', 'VWD <21d (Mittel)', 'VWD <21d (Med)']
            for c, h in enumerate(vwd_h): ws.write(curr_r+1, c, h, f_header)
            for i, y in enumerate([2024, 2025, 2026]):
                y_df = df[df['Year'] == y]
                v_short = y_df[(y_df['VWD_num'] > 0) & (y_df['VWD_num'] < 21)]['VWD_num']
                ws.write(curr_r+2+i, 0, y, f_cell)
                ws.write(curr_r+2+i, 1, float(y_df[y_df['VWD_num'] > 0]['VWD_num'].median()) if not y_df.empty else 0, f_cell)
                ws.write(curr_r+2+i, 2, float(v_short.mean()) if not v_short.empty else 0, f_num)
                ms = float(v_short.median()) if not v_short.empty else 0
                ws.write(curr_r+2+i, 3, ms, f_green if 0 < ms <= 5 else f_cell)

            # --- 3. SPRECHSTUNDE ---
            curr_r = 17
            ws.write(curr_r, 0, '3. ZUWEISUNG ÜBER KLAPPENSPRECHSTUNDE', f_title)
            df['KS_bool'] = df['KS'].apply(lambda x: 1 if str(x).lower() in ['x', '1', 'ja'] else 0)
            for i, y in enumerate([2025, 2026]):
                ks_c = int(df[df['Year'] == y]['KS_bool'].sum())
                ws.write(curr_r+1+i, 0, y, f_cell)
                ws.write(curr_r+1+i, 1, ks_c, f_cell)

            # --- 4. STRATEGIE ---
            curr_r = 21
            ws.write(curr_r, 0, '4. STRATEGIE: DEVICE-MIX 2026', f_title)
            tavi_26 = df_2026[df_2026['KPI_Kat'] == 'TAVI']
            ev_r = float(tavi_26['Device'].str.contains('Evolut', na=False, case=False).mean()) if len(tavi_26) > 0 else 0
            ws.write(curr_r+1, 0, 'Evolut-Anteil (TAVI) - Ziel 80%', f_cell)
            ws.write(curr_r+1, 1, ev_r, f_pct)

            teer_26 = df_2026[df_2026['KPI_Kat'].isin(['MTEER', 'TTEER'])]
            pascal = int(teer_26['Device'].str.contains('Pascal', na=False, case=False).sum())
            clip = int(teer_26['Device'].str.contains('Clip', na=False, case=False).sum())
            p_ratio = pascal / (pascal + clip) if (pascal + clip) > 0 else 0
            ws.write(curr_r+2, 0, 'Pascal-Anteil (TEER: TriClip/MitraClip)', f_cell)
            ws.write(curr_r+2, 1, p_ratio, f_pct)

            # --- 5. TEAMS ---
            curr_r = 28
            ws.write(curr_r, 0, '5. TAVI-TEAMS 2026', f_title)
            t_stats = tavi_26['Team'].value_counts().reset_index()
            for c, h in enumerate(['Team', 'Fälle', 'Anteil']): ws.write(curr_r+1, c, h, f_header)
            for i, row in enumerate(t_stats.values):
                ws.write(curr_r+2+i, 0, str(row[0]), f_cell)
                ws.write(curr_r+2+i, 1, int(row[1]), f_cell)
                ws.write(curr_r+2+i, 2, float(row[1]/len(tavi_26)) if len(tavi_26) > 0 else 0, f_pct)

            # --- 6. QUALITÄT ---
            curr_r = 37
            ws.write(curr_r, 0, '6. QUALITÄT & KOMPLIKATIONSRATEN 2026', f_title)
            comp_dict = {'Tod w. Aufenth.': ['Mortalität', 0.02], 'Stroke': ['Apoplex', 0.015], 'SM_neu': ['Schrittmacher', 0.10], 'Gefäß_Kom.': ['Gefäßkompl.', 0.05]}
            for c, h in enumerate(['Indikator', 'Fälle', 'Rate %', 'Benchmark']): ws.write(curr_r+1, c, h, f_header)
            for i, (col, lb) in enumerate(comp_dict.items()):
                val = int(pd.to_numeric(df_2026[col], errors='coerce').fillna(0).sum())
                rate = val / len(df_2026) if len(df_2026) > 0 else 0
                ws.write(curr_r+2+i, 0, lb[0], f_cell)
                ws.write(curr_r+2+i, 1, val, f_cell)
                ws.write(curr_r+2+i, 2, float(rate), f_green if rate <= lb[1] else f_red)
                ws.write(curr_r+2+i, 3, float(lb[1]), f_pct)

            # --- 7. VERLAUF ---
            curr_r = 46
            ws.write(curr_r, 0, '7. VERLAUF DER FALLZAHLEN 2021 - 2026', f_title)
            years = [2021, 2022, 2023, 2024, 2025, 2026]
            for c, h in enumerate(['Jahr', 'TAVI', 'TEER', 'Gesamt']): ws.write(curr_r+1, c, h, f_header)
            for i, y in enumerate(years):
                y_df = df[df['Year'] == y]
                ws.write(curr_r+2+i, 0, y, f_cell)
                ws.write(curr_r+2+i, 1, int(len(y_df[y_df['KPI_Kat'] == 'TAVI'])), f_cell)
                ws.write(curr_r+2+i, 2, int(len(y_df[y_df['KPI_Kat'].isin(['MTEER', 'TTEER'])])), f_cell)
                ws.write(curr_r+2+i, 3, int(len(y_df)), f_cell)
            
            chart = workbook.add_chart({'type': 'line'})
            for col in range(1, 4):
                chart.add_series({
                    'name':       ['Master Dashboard', curr_r+1, col],
                    'categories': ['Master Dashboard', curr_r+2, 0, curr_r+7, 0],
                    'values':     ['Master Dashboard', curr_r+2, col, curr_r+7, col],
                })
            ws.insert_chart(curr_r+2, 5, chart)

            ws.set_column('A:A', 35); ws.set_column('B:Q', 12)

        st.success("✅ Dashboard generiert!")
        st.download_button("📊 Download Excel", output.getvalue(), f"Herzklappen_{heute_str}.xlsx")

    except Exception as e:
        st.error(f"❌ Fehler: {e}")
