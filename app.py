import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import datetime
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Dialyse Dienstplan", layout="wide")
st.title("🏥 Dienstplan-Generator: Dialyse (Stunden-Optimiert)")
st.markdown("Laden Sie Ihre Excel-Datei hoch. Die KI erstellt den Plan, achtet strikt auf die +/- 25h Grenze und verteilt die Stunden fair.")

uploaded_file = st.file_uploader("Excel-Tabelle hochladen (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_haupt = pd.read_excel(uploaded_file, sheet_name=0)
        df_haupt.fillna("Leer", inplace=True)
        
        try:
            df_regeln = pd.read_excel(uploaded_file, sheet_name=1)
            df_regeln.fillna("", inplace=True)
        except:
            st.error("Fehler: Konnte das zweite Tabellenblatt 'Sonderregeln' nicht finden. Bitte legen Sie es an.")
            st.stop()
            
        st.success("Tabelle und Sonderregeln erfolgreich eingelesen!")
        
        if st.button("Dienstplan jetzt berechnen"):
            with st.spinner("Die KI balanciert nun die Stundenkonten aus... Dies kann einen Moment dauern."):
                
                model = cp_model.CpModel()
                
                mitarbeiter_liste = df_haupt['Name'].tolist()
                tage = [col for col in df_haupt.columns if col not in ['Name', 'Stundenausmaß', 'Berechnung Soll-Arbeitszeit', 'Übertrag Vormonat']]
                num_tage = len(tage)
                schichten = ['D1', 'V1', 'SL', 'D7', 'Frei', 'U', 'ZB', 'ÜZA', 'BA', 'FB']
                
                dienst_vars = {}
                for m in mitarbeiter_liste:
                    for t_idx in range(num_tage):
                        for s in schichten:
                            dienst_vars[(m, t_idx, s)] = model.NewBoolVar(f"{m}_{t_idx}_{s}")
                
                # --- 0. ZAHLEN SICHER EINLESEN ---
                def parse_num(val):
                    try:
                        if pd.isna(val) or str(val).strip() == "Leer" or str(val).strip() == "": return 0.0
                        return float(str(val).replace(',', '.').replace('+', '').strip())
                    except:
                        return 0.0

                ma_daten = {}
                for index, row in df_haupt.iterrows():
                    m = row['Name']
                    ma_daten[m] = {
                        'ausmass_int': int(parse_num(row.get('Stundenausmaß', 0)) * 100),
                        'soll_int': int(parse_num(row.get('Berechnung Soll-Arbeitszeit', 0)) * 100),
                        'uebertrag_int': int(parse_num(row.get('Übertrag Vormonat', 0)) * 100)
                    }

                # --- 1. ROBUSTE WOCHENTAGS-ERKENNUNG ---
                wochentage_idx = []
                samstage_idx = []
                for t_idx, tag_val in enumerate(tage):
                    if isinstance(tag_val, datetime.datetime):
                        wt = tag_val.weekday()
                    else:
                        try:
                            dt = pd.to_datetime(str(tag_val), dayfirst=True)
                            wt = dt.weekday()
                        except:
                            wt = t_idx % 7 
                            
                    wochentage_idx.append(wt)
                    if wt == 5: 
                        samstage_idx.append(t_idx)

                # --- 2. GRUNDREGELN & EXCEL-INPUT ---
                feste_eintraege = {} 
                
                for m in mitarbeiter_liste:
                    for t_idx in range(num_tage):
                        model.AddExactlyOne([dienst_vars[(m, t_idx, s)] for s in schichten])
                        
                for index, row in df_haupt.iterrows():
                    m = row['Name']
                    for t_idx, tag in enumerate(tage):
                        if wochentage_idx[t_idx] == 6: 
                            continue
                            
                        eintrag = str(row[tag]).strip()
                        if eintrag != "Leer":
                            feste_eintraege[(m, t_idx)] = eintrag 
                            if eintrag == "F": 
                                model.Add(dienst_vars[(m, t_idx, 'Frei')] == 1)
                            elif eintrag in schichten:
                                model.Add(dienst_vars[(m, t_idx, eintrag)] ==
