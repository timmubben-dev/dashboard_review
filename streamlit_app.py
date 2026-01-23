import streamlit as st
import pandas as pd
import numpy as np
import io
import msoffcrypto
import altair as alt
from datetime import datetime

# 1. SEITEN-KONFIGURATION
st.set_page_config(page_title="Herzklappen Master-Dashboard", layout="wide")
st.title("🏥 Herzklappen Master-Dashboard 2026")

# Sidebar für Login & Upload
st.sidebar.header("🔐 Datensicherheit")
password = st.sidebar.text_input("Bitte Excel-Passwort eingeben", type="password")
uploaded_file = st.sidebar.file_uploader("Verschlüsselte Excel-Datei wählen", type=["xlsx"])

def map_to_kpi(e):
    e = str(e).lower()
    if 'tavi' in e: return 'TAVI'
    if 'edge-to-edge mk' in e or 'tmvi' in e: return 'MTEER'
    if 'edge-to-edge tk' in e or 'htp tk' in e: return 'TTEER'
    if 'ttvi' in e: return 'TTVI'
    if 'tricvalve' in e or 'ttvr' in e: return 'TTVR'
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
        
        # --- VISUALISIERUNG IN STREAMLIT ---
        st.subheader("📈 Langzeit-Trend der Fallzahlen (2022-2026)")
        h_cats = ['TAVI', 'MTEER', 'TTEER']
        trend_data = df[df['Year'].between(2022, 2026)].groupby(['Year', 'KPI_Kat']).size().reset_index(name='Anzahl')
        trend_data = trend_data[trend_data['KPI_Kat'].isin(h_cats)]
        
        trend_chart = alt.Chart(trend_data).mark_line(point=True).encode(
            x=alt.X('Year:O', title='Jahr'),
            y=alt.Y('Anzahl:Q', title='Fallzahl'),
            color=alt.Color('KPI_Kat:N', title='Eingriff'),
            tooltip=['Year', 'KPI_Kat', 'Anzahl']
        ).properties(height=350).interactive()
        st.altair_chart(trend_chart, use_container_width=True)

        # --- EXCEL GENERIERUNG ---
        df_2026 = df[df['Year'] == 2026].copy()
        months_passed = df_2026['Month'].max() or 1
        heute_str = datetime.now().strftime('%d-%m-%Y')

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            ws = workbook.add_worksheet('Master Dashboard')
            
            # Formate
            title_f = workbook.add_format({'bold': True, 'size': 12, 'font_color': '#1F4E78', 'bottom': 2})
            header_f = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
            cell_f = workbook.add_format({'border': 1, 'align': 'center'})
            pct_f = workbook.add_format({'num_format': '0.0%', 'border': 1, 'align': 'center'})
            red_f = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center'})
            num_f = workbook.add_format({'num_format': '0.0', 'border': 1, 'align': 'center'})

            # --- PUNKT 1: LEISTUNGSZAHLEN ---
            ws.write('A1', '1. LEISTUNGSZAHLEN & PROGNOSE 2026', title_f)
            targets = {'TAVI': 46, 'MTEER': 10, 'TTEER': 7, 'TTVI': 3, 'TTVR': 1}
            headers = ['Kategorie'] + ['Jan','Feb','Mrz','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Dez'] + ['YTD','Prognose','Soll','Status']
            for c, h in enumerate(headers): ws.write(2, c, h, header_f)
            
            for r, (cat, t_mo) in enumerate(targets.items()):
                c_df = df_2026[df_2026['KPI_Kat'] == cat]
                counts = c_df.groupby('Month').size()
                ws.write(r+3, 0, cat, workbook.add_format({'bold': True, 'border': 1}))
                for m in range(1, 13):
                    val = counts.get(m, 0)
                    fmt = red_f if (val < t_mo and m <= months_passed) else cell_f
                    ws.write(r+3, m, val, fmt)
                ist_y = len(c_df); prog = round((ist_y / months_passed) * 12)
                ws.write(r+3, 13, ist_y, cell_f); ws.write(r+3, 14, prog, cell_f)
                ws.write(r+3, 15, t_mo * 12, cell_f)
                ws.write_formula(r+3, 16, f'=IFERROR(O{r+4}/P{r+4}, 0)', pct_f)

            # --- PUNKT 2: VERWEILDAUER ---
            curr_r = 10
            ws.write(curr_r, 0, '2. VERWEILDAUER (ZIEL: 5 TAGE MEDIAN)', title_f)
            v_headers = ['Jahr', 'VWD Alle (Med)', 'VWD <21d (Mittel)', 'VWD <21d (Med)']
            for c, h in enumerate(v_headers): ws.write(curr_r+1, c, h, header_f)
            for i, y in enumerate([2024, 2025, 2026]):
                y_df = df[df['Year'] == y]
                v_short = y_df[(y_df['VWD_num'] > 0) & (y_df['VWD_num'] < 21)]['VWD_num']
                ws.write(curr_r+2+i, 0, y, cell_f)
                ws.write(curr_r+2+i, 1, y_df[y_df['VWD_num'] > 0]['VWD_num'].median() if not y_df.empty else 0, cell_f)
                ws.write(curr_r+2+i, 2, v_short.mean() if not v_short.empty else 0, num_f)
                ws.write(curr_r+2+i, 3, v_short.median() if not v_short.empty else 0, cell_f)

            # --- PUNKT 3: SPRECHSTUNDE ---
            curr_r = 17
            ws.write(curr_r, 0, '3. ZUWEISUNG ÜBER KLAPPENSPRECHSTUNDE', title_f)
            df['KS_bool'] = df['KS'].apply(lambda x: 1 if str(x).lower() in ['x', '1', 'ja'] else 0)
            for i, y in enumerate([2025, 2026]):
                ks_c = df[df['Year'] == y]['KS_bool'].sum()
                ws.write(curr_r+1+i, 0, y, cell_f); ws.write(curr_r+1+i, 1, ks_c, cell_f)

            # --- PUNKT 4: TEAMS ---
            curr_r = 21
            ws.write(curr_r, 0, '4. TAVI-TEAMS 2026', title_f)
            tavi_26 = df_2026[df_2026['KPI_Kat'] == 'TAVI']
            t_stats = tavi_26['Team'].value_counts().reset_index()
            for c, h in enumerate(['Team', 'Fälle', 'Anteil']): ws.write(curr_r+1, c, h, header_f)
            for i, row in enumerate(t_stats.values):
                ws.write(curr_r+2+i, 0, row[0], cell_f); ws.write(curr_r+2+i, 1, row[1], cell_f)
                ws.write(curr_r+2+i, 2, row[1]/len(tavi_26) if len(tavi_26) > 0 else 0, pct_f)

            # --- PUNKT 5: STRATEGIE ---
            curr_r = 30
            ws.write(curr_r, 0, '5. STRATEGIE & QUALITÄT 2026', title_f)
            ev_r = tavi_26['Device'].str.contains('Evolut', na=False, case=False).mean() if len(tavi_26) > 0 else 0
            ws.write(curr_r+1, 0, 'Evolut-Anteil (Ziel 80%)', cell_f); ws.write(curr_r+1, 1, ev_r, pct_f)

            # --- PUNKT 6: HISTORIE ---
            curr_r = 34
            ws.write(curr_r, 0, '6. HISTORISCHE ENTWICKLUNG', title_f)
            h_cats_list = ['TAVI', 'MTEER', 'TTEER']
            for c, h in enumerate(['Jahr'] + h_cats_list): ws.write(curr_r+1, c, h, header_f)
            for i, y in enumerate([2022, 2023, 2024, 2025, 2026]):
                ws.write(curr_r+2+i, 0, y, cell_f)
                for j, cat in enumerate(h_cats_list):
                    ws.write(curr_r+2+i, j+1, len(df[(df['Year'] == y) & (df['KPI_Kat'] == cat)]), cell_f)

            # --- PUNKT 7: KOMPLIKATIONEN ---
            curr_r = 42
            ws.write(curr_r, 0, '7. KOMPLIKATIONSRATEN 2026', title_f)
            comp_dict = {'Tod w. Aufenth.': 'Mortalität', 'Stroke': 'Apoplex', 'SM_neu': 'Schrittmacher', 'Gefäß_Kom.': 'Gefäßkompl.'}
            for c, h in enumerate(['Indikator', 'Fälle', 'Rate %']): ws.write(curr_r+1, c, h, header_f)
            for i, (col, lab) in enumerate(comp_dict.items()):
                val = pd.to_numeric(df_2026[col], errors='coerce').fillna(0).sum()
                rate = val / len(df_2026) if len(df_2026) > 0 else 0
                ws.write(curr_r+2+i, 0, lab, cell_f); ws.write(curr_r+2+i, 1, val, cell_f); ws.write(curr_r+2+i, 2, rate, pct_f)

            ws.set_column('A:A', 35); ws.set_column('B:Q', 12)

        st.success("✅ Dashboard vollständig generiert!")
        st.download_button(label="📊 Dashboard herunterladen", data=output.getvalue(), file_name=f"Master_Dashboard_{heute_str}.xlsx")

    except Exception as e:
        st.error(f"❌ Fehler: {e}. Bitte Passwort prüfen.")
