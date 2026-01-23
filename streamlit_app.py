import streamlit as st
import pandas as pd
import numpy as np
import io
import msoffcrypto
import altair as alt
from datetime import datetime

# 1. SEITEN-KONFIGURATION
st.set_page_config(page_title="Herzklappen Master-Dashboard", layout="wide")
st.title("🏥 Herzklappen Master-Dashboard Generator 2026")

# Sidebar für Login & Upload
st.sidebar.header("🔐 Datensicherheit")
# Hier fragt er aktiv nach dem Passwort
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

if uploaded_file:
    if not password:
        st.warning("⚠️ Bitte geben Sie zuerst das Passwort in der Seitenleiste ein.")
    else:
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
            
            # --- VISUALISIERUNG (Trend) ---
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

            # --- EXCEL MASTER DASHBOARD GENERIERUNG ---
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
                green_f = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center'})

                # (Punkte 1-6 bleiben wie gehabt...)
                # 1. Leistungszahlen
                ws.write('A1', '1. LEISTUNGSZAHLEN & PROGNOSE 2026', title_f)
                targets = {'TAVI': 46, 'MTEER': 10, 'TTEER': 7, 'TTVI': 3, 'TTVR': 1}
                for r, (cat, t_mo) in enumerate(targets.items()):
                    cat_df = df_2026[df_2026['KPI_Kat'] == cat]
                    ist_ytd = len(cat_df)
                    ws.write(r+3, 0, cat, cell_f)
                    ws.write(r+3, 13, ist_ytd, cell_f)
                    # ... (weitere Details wie im Vorcode)

                # --- NEU: PUNKT 7: KOMPLIKATIONSRATEN ---
                ws.write('A50', '7. KOMPLIKATIONSRATEN 2026 (BENCHMARK-VERGLEICH)', title_f)
                comp_headers = ['Indikator', 'Fälle (n)', 'Rate (%)', 'Benchmark']
                for c, h in enumerate(comp_headers): ws.write(51, c, h, header_f)
                
                complications = {
                    'SM_neu': ['Schrittmacher-Pflicht', 0.10],
                    'Tod w. Aufenth.': ['In-Hospital Mortalität', 0.02],
                    'Stroke': ['Apoplex-Rate', 0.015],
                    'Gefäß_Kom.': ['Schwere Gefäßkompl.', 0.05]
                }
                
                for r, (col, details) in enumerate(complications.items()):
                    val_sum = pd.to_numeric(df_2026[col], errors='coerce').fillna(0).sum()
                    rate = val_sum / len(df_2026) if len(df_2026) > 0 else 0
                    ws.write(52+r, 0, details[0], cell_f)
                    ws.write(52+r, 1, val_sum, cell_f)
                    ws.write(52+r, 2, rate, pct_f)
                    ws.write(52+r, 3, details[1], pct_f)

                ws.set_column('A:A', 36)

            st.success("✅ Dashboard erfolgreich generiert.")
            st.download_button(label="📥 Master-Dashboard herunterladen", 
                               data=output.getvalue(), 
                               file_name=f"Herzklappen_Master_{heute_str}.xlsx")

        except Exception as e:
            st.error(f"❌ Fehler: {e}. Prüfen Sie Passwort und Format.")
